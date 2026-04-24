from flask import Flask, request, jsonify, session, redirect, url_for, send_from_directory
from flask_cors import CORS
from werkzeug.middleware.proxy_fix import ProxyFix
import os
import sys
import json
import re
import uuid

# Add parent directory to path for imports
_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(_BASE_DIR)
FRONTEND_DIR = os.path.join(_BASE_DIR, 'frontend')

from config import Config
from backend.session_manager import SessionManager

config = Config()
IS_PRODUCTION = config.FLASK_ENV == 'production'

app = Flask(__name__)
app.config['SECRET_KEY'] = config.SECRET_KEY
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
# Require HTTPS for the session cookie in production so it can't leak over plain HTTP.
app.config['SESSION_COOKIE_SECURE'] = IS_PRODUCTION
app.config['SESSION_COOKIE_HTTPONLY'] = True

# When behind a reverse proxy (Railway/Render/nginx/Caddy), trust the X-Forwarded-*
# headers so Flask knows the real scheme (https) and host. This is required for
# Google OAuth redirect URIs to be generated with https:// when terminating TLS
# at the proxy layer.
if IS_PRODUCTION:
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

CORS(app, supports_credentials=True)  # Enable CORS for frontend with credentials

# Store OAuth credentials in memory (in production, use Redis or database)
oauth_credentials_store = {}

# Per-user sessions: each user gets their own sheets_service and automation_service
session_manager = SessionManager(config, max_sessions=10, idle_timeout_seconds=1800)


def _get_user_session():
    """Get (or create) the UserSession for the current Flask session.

    Raises if server is at capacity.
    """
    if 'sid' not in session:
        session['sid'] = str(uuid.uuid4())
        session.permanent = False
    return session_manager.get_or_create(session['sid'])


def authenticate_sheets_service(sheets_service, data):
    """Authenticate the given sheets_service using credentials in request data."""
    oauth_state = data.get('oauth_state')
    service_account_file = data.get('service_account_file')
    service_account_json = data.get('service_account_json')

    if oauth_state:
        if oauth_state not in oauth_credentials_store or 'credentials' not in oauth_credentials_store[oauth_state]:
            raise Exception('OAuth credentials not found. Please authorize again.')
        creds_dict = oauth_credentials_store[oauth_state]['credentials']
        sheets_service.authenticate_with_oauth(creds_dict)
    elif service_account_file or service_account_json:
        sheets_service.authenticate(service_account_file, service_account_json)
    elif not sheets_service.gc:
        try:
            sheets_service.authenticate()
        except Exception:
            raise Exception('Google authentication required. Please use OAuth or provide service account credentials.')

@app.route('/api/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok'})

@app.route('/api/oauth/config', methods=['GET'])
def get_oauth_config():
    """Get OAuth configuration for debugging"""
    return jsonify({
        'client_id': config.GOOGLE_CLIENT_ID[:30] + '...' if config.GOOGLE_CLIENT_ID else None,
        'redirect_uri': config.GOOGLE_REDIRECT_URI,
        'instructions': 'Make sure these URIs are added to Google Cloud Console',
        'required_javascript_origins': [
            'http://localhost:8000',
            'http://localhost:5001',
            'http://127.0.0.1:5500'
        ],
        'required_redirect_uris': [
            config.GOOGLE_REDIRECT_URI
        ],
        'console_url': 'https://console.cloud.google.com/apis/credentials'
    })

@app.route('/api/sheets/oauth/client-id', methods=['GET'])
def get_client_id():
    """Get Google OAuth Client ID for frontend"""
    client_id = config.GOOGLE_CLIENT_ID
    if not client_id:
        return jsonify({'client_id': None, 'error': 'GOOGLE_CLIENT_ID not configured'})
    return jsonify({'client_id': client_id})

@app.route('/api/sheets/oauth/exchange-token', methods=['POST'])
def exchange_token():
    """Exchange Google JWT credential for access token"""
    try:
        from google.auth.transport import requests
        from google.oauth2 import id_token
        import uuid
        
        data = request.json
        credential = data.get('credential')
        frontend_url = data.get('frontend_url')  # Get frontend URL from request

        if not credential:
            return jsonify({'success': False, 'error': 'Credential required'}), 400

        if not config.GOOGLE_CLIENT_ID or not config.GOOGLE_CLIENT_SECRET:
            return jsonify({'success': False, 'error': 'OAuth is not configured on the server.'}), 500

        # Verify the JWT token
        try:
            idinfo = id_token.verify_oauth2_token(
                credential,
                requests.Request(),
                config.GOOGLE_CLIENT_ID
            )
        except ValueError as e:
            return jsonify({'success': False, 'error': f'Invalid token: {str(e)}'}), 400
        
        # Get user info
        user_email = idinfo.get('email')
        
        # Now we need to get an access token for Google Sheets API
        # We'll use the authorization code flow instead
        # For now, store the credential and use it to get access token via OAuth flow
        
        # Generate a state for this session
        state = str(uuid.uuid4())
        
        # Store the credential temporarily
        oauth_credentials_store[state] = {
            'credential': credential,
            'user_email': user_email,
            'idinfo': idinfo
        }
        
        # We need to initiate OAuth flow to get access token
        # But we can use the credential to get user info and then request access
        from google_auth_oauthlib.flow import Flow
        
        print(f"=== OAuth Exchange Token ===")
        print(f"GOOGLE_CLIENT_ID: {config.GOOGLE_CLIENT_ID[:20]}..." if config.GOOGLE_CLIENT_ID else None)
        print(f"GOOGLE_REDIRECT_URI: {config.GOOGLE_REDIRECT_URI}")

        scopes = [
            'openid',
            'https://www.googleapis.com/auth/userinfo.email',
            'https://www.googleapis.com/auth/userinfo.profile',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ]

        client_config = {
            "web": {
                "client_id": config.GOOGLE_CLIENT_ID,
                "client_secret": config.GOOGLE_CLIENT_SECRET,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [config.GOOGLE_REDIRECT_URI]
            }
        }

        flow = Flow.from_client_config(
            client_config,
            scopes=scopes,
            redirect_uri=config.GOOGLE_REDIRECT_URI
        )
        
        # Generate authorization URL - flow creates its own state
        authorization_url, flow_state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true',
            prompt='consent',
            login_hint=user_email  # Pre-fill the email
        )
        
        print(f"Generated authorization URL: {authorization_url[:100]}...")
        print(f"Flow state: {flow_state}")
        
        # Get frontend URL from request body, headers, or use default
        if not frontend_url:
            frontend_url = request.headers.get('Origin') or request.headers.get('Referer', 'http://localhost:8000/index.html')
        
        # Clean up the frontend URL - remove query params and hash, but keep the path
        try:
            from urllib.parse import urlparse, urlunparse
            parsed = urlparse(frontend_url)
            # Normalize path - if it's just '/', use '/index.html'
            path = parsed.path if parsed.path and parsed.path != '/' else '/index.html'
            # Reconstruct URL without query or fragment
            frontend_url = urlunparse((
                parsed.scheme,
                parsed.netloc,
                path,
                '',  # params
                '',  # query
                ''   # fragment
            ))
        except Exception as e:
            # Fallback to default
            print(f"Error parsing frontend URL '{frontend_url}': {e}")
            frontend_url = 'http://localhost:8000/index.html'
        
        print(f"Storing frontend URL: {frontend_url}")
        
        # IMPORTANT: Store using flow_state (the state Google will send back)
        # not our custom UUID
        oauth_credentials_store[flow_state] = {
            'credential': credential,
            'user_email': user_email,
            'idinfo': idinfo,
            'flow': flow,
            'frontend_url': frontend_url
        }
        
        # Return authorization URL for frontend to open
        return jsonify({
            'success': True,
            'authorization_url': authorization_url,
            'state': flow_state,  # Use the flow's state, not a custom UUID
            'message': 'Please complete authorization'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/sheets/oauth/authorize', methods=['POST'])
def oauth_authorize():
    """Start OAuth2 flow - returns authorization URL"""
    try:
        from google_auth_oauthlib.flow import Flow
        
        # Get OAuth credentials from request or config
        data = request.json or {}
        client_id = data.get('client_id') or config.GOOGLE_CLIENT_ID
        client_secret = data.get('client_secret') or config.GOOGLE_CLIENT_SECRET
        redirect_uri = data.get('redirect_uri') or config.GOOGLE_REDIRECT_URI
        
        if not client_id or not client_secret:
            return jsonify({
                'success': False,
                'error': 'OAuth credentials required. Please provide client_id and client_secret, or set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET in .env'
            }), 400
        
        # OAuth scopes
        scopes = [
            'openid',
            'https://www.googleapis.com/auth/userinfo.email',
            'https://www.googleapis.com/auth/userinfo.profile',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive.readonly'
        ]
        
        # Create OAuth flow
        client_config = {
            "web": {
                "client_id": client_id,
                "client_secret": client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": [redirect_uri]
            }
        }
        
        flow = Flow.from_client_config(
            client_config,
            scopes=scopes,
            redirect_uri=redirect_uri
        )
        
        # Generate authorization URL
        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true',
            prompt='consent'
        )
        
        # Store flow in session (using state as key)
        oauth_credentials_store[state] = {
            'flow': flow,
            'client_id': client_id,
            'client_secret': client_secret,
            'redirect_uri': redirect_uri
        }
        
        return jsonify({
            'success': True,
            'authorization_url': authorization_url,
            'state': state
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/sheets/oauth/callback', methods=['GET'])
def oauth_callback():
    """Handle OAuth2 callback"""
    print(f"=== OAuth Callback Hit ===")
    print(f"Request URL: {request.url}")
    print(f"Request args: {request.args}")
    try:
        from google_auth_oauthlib.flow import Flow
        
        state = request.args.get('state')
        code = request.args.get('code')
        
        print(f"State: {state}")
        print(f"Code: {code[:20] if code else None}...")
        print(f"Available states in store: {list(oauth_credentials_store.keys())}")
        
        if not state or not code:
            return jsonify({'success': False, 'error': 'Missing state or code'}), 400
        
        if state not in oauth_credentials_store:
            return jsonify({'success': False, 'error': 'Invalid state. Please start authorization again.'}), 400
        
        stored = oauth_credentials_store[state]
        flow = stored['flow']
        
        # Exchange code for tokens
        flow.fetch_token(code=code)
        credentials = flow.credentials
        
        # Store credentials (convert to dict for JSON serialization)
        creds_dict = {
            'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes
        }
        
        # Store credentials with state
        oauth_credentials_store[state]['credentials'] = creds_dict
        
        # Get frontend URL from stored state or use default
        stored_data = oauth_credentials_store.get(state, {})
        frontend_url = stored_data.get('frontend_url', 'http://localhost:8000/index.html')
        
        print(f"=== OAuth Callback Redirect ===")
        print(f"State: {state}")
        print(f"Frontend URL to redirect to: {frontend_url}")
        print(f"Available states: {list(oauth_credentials_store.keys())}")
        if state in oauth_credentials_store:
            print(f"Stored data keys: {list(oauth_credentials_store[state].keys())}")
        
        # Return HTML page that automatically redirects back
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Authorization Successful</title>
        </head>
        <body>
            <h2>Authorization Successful!</h2>
            <p>Redirecting back to application...</p>
            <script>
                // Send message to parent window if in popup
                if (window.opener) {{
                    window.opener.postMessage({{
                        type: 'oauth_success',
                        state: '{state}'
                    }}, '*');
                    // Close popup immediately
                    window.close();
                }} else {{
                    // If not in popup, redirect back to frontend immediately
                    const redirectUrl = '{frontend_url}?oauth_success=true&state={state}';
                    console.log('Redirecting to:', redirectUrl);
                    window.location.href = redirectUrl;
                }}
            </script>
        </body>
        </html>
        '''
    except Exception as e:
        error_msg = str(e).replace("'", "&#39;").replace('"', '&quot;')  # HTML escape
        
        # Try to get frontend URL from state if available
        state = request.args.get('state')
        frontend_url = 'http://localhost:8000/index.html'  # Default for port 8000
        if state and state in oauth_credentials_store:
            frontend_url = oauth_credentials_store[state].get('frontend_url', frontend_url)
        
        return f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Authorization Failed</title>
        </head>
        <body>
            <h2>Authorization Failed</h2>
            <p>Error: {error_msg}</p>
            <p>Redirecting back to application...</p>
            <script>
                if (window.opener) {{
                    window.opener.postMessage({{
                        type: 'oauth_error',
                        error: '{error_msg}'
                    }}, '*');
                    // Close popup immediately
                    window.close();
                }} else {{
                    // If not in popup, redirect back to frontend with error
                    window.location.href = '{frontend_url}?oauth_error=true&error=' + encodeURIComponent('{error_msg}');
                }}
            </script>
        </body>
        </html>
        ''', 400

@app.route('/api/sheets/oauth/use-credentials', methods=['POST'])
def use_oauth_credentials():
    """Use stored OAuth credentials to authenticate"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service

        data = request.json
        state = data.get('state')

        if not state or state not in oauth_credentials_store:
            return jsonify({'success': False, 'error': 'Invalid state. Please authorize again.'}), 400

        stored = oauth_credentials_store[state]
        if 'credentials' not in stored:
            return jsonify({'success': False, 'error': 'Credentials not found. Please authorize again.'}), 400

        # Use credentials to authenticate
        creds_dict = stored['credentials']
        sheets_service.authenticate_with_oauth(creds_dict)

        return jsonify({'success': True, 'message': 'Authenticated successfully'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/sheets/connect', methods=['POST'])
def connect_sheets():
    """Connect to Google Sheets"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service

        data = request.json
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        service_account_file = data.get('service_account_file')
        service_account_json = data.get('service_account_json')
        oauth_state = data.get('oauth_state')  # For OAuth flow
        
        print(f"=== Connect Sheets Request ===")
        print(f"Sheet URL: {sheet_url}")
        print(f"Sheet name: {sheet_name}")
        print(f"OAuth state: {oauth_state}")
        print(f"Available states: {list(oauth_credentials_store.keys())}")
        
        if not sheet_url:
            return jsonify({'success': False, 'error': 'Sheet URL is required'}), 400
        if not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet name is required'}), 400
        
        # Authenticate if OAuth state provided
        if oauth_state:
            print(f"Using OAuth state: {oauth_state}")
            if oauth_state not in oauth_credentials_store:
                print(f"ERROR: OAuth state not found in store")
                return jsonify({'success': False, 'error': f'OAuth state not found. Available states: {list(oauth_credentials_store.keys())}'}, 400)
            if 'credentials' not in oauth_credentials_store[oauth_state]:
                print(f"ERROR: Credentials not found for state")
                return jsonify({'success': False, 'error': 'OAuth credentials not found. Please authorize again.'}), 400
            creds_dict = oauth_credentials_store[oauth_state]['credentials']
            print(f"Authenticating with OAuth credentials")
            sheets_service.authenticate_with_oauth(creds_dict)
        # Authenticate if service account provided
        elif service_account_file or service_account_json:
            print(f"Using service account")
            sheets_service.authenticate(service_account_file, service_account_json)
        elif not sheets_service.gc:
            # Try to authenticate with config
            print(f"Trying to authenticate with config")
            try:
                sheets_service.authenticate()
            except Exception as auth_error:
                print(f"ERROR: Authentication failed: {str(auth_error)}")
                return jsonify({
                    'success': False, 
                    'error': 'Google authentication required. Please use OAuth or provide service account credentials.'
                }), 400
        
        print(f"Connecting to sheet...")
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        print(f"Getting columns...")
        columns = sheets_service.get_columns(worksheet)
        print(f"Success! Found {len(columns)} columns")
        
        return jsonify({
            'success': True,
            'columns_count': len(columns),
            'columns': columns
        })
    except Exception as e:
        print(f"ERROR in connect_sheets: {str(e)}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/login', methods=['POST'])
def login():
    """Login to myaccount.brown.edu"""
    try:
        user_session = _get_user_session()
        automation_service = user_session.automation_service

        data = request.json
        username = data.get('username') or config.MYACCOUNT_USERNAME
        password = data.get('password') or config.MYACCOUNT_PASSWORD
        
        if not username or not password:
            return jsonify({'success': False, 'error': 'Username and password required'}), 400
        
        # Check if driver is alive, cleanup if not
        if automation_service.driver and not automation_service._is_driver_alive():
            print("Detected dead driver session, cleaning up...")
            automation_service._cleanup_driver()
        
        success = automation_service.login(username, password)
        
        if success:
            return jsonify({'success': True, 'message': 'Login successful. Please approve Duo push.'})
        else:
            return jsonify({'success': False, 'error': 'Login failed'}), 401
    except Exception as e:
        error_msg = str(e)
        print(f"Login error: {error_msg}")

        # If it's a driver connection error, cleanup and provide better error message
        if "disconnected" in error_msg.lower() or "unable to connect" in error_msg.lower():
            try:
                user_session.automation_service._cleanup_driver()
            except Exception:
                pass
            return jsonify({
                'success': False,
                'error': 'Browser connection lost. Please try logging in again.'
            }), 500

        return jsonify({'success': False, 'error': error_msg}), 500

@app.route('/api/automation/abort', methods=['POST'])
def abort_automation():
    """Abort the current automation task. Navigator returns to myaccount."""
    try:
        user_session = _get_user_session()
        automation_service = user_session.automation_service
        automation_service.abort()
        return jsonify({'success': True, 'message': 'Abort requested. Task will stop after current item.'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/add', methods=['POST'])
def add_privileges():
    """Add privileges to users - reads from Google Sheets column E (index 4)"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service
        automation_service = user_session.automation_service

        data = request.json
        print(f"📋 Received request data: {data}")  # Debug logging

        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        app_name = data.get('app_name', 'Slate [SLATE]')
        comment = data.get('comment', '')
        performed_by_name = data.get('performed_by_name')
        column_index = data.get('column_index', 4)  # Column E (index 4)
        oauth_state = data.get('oauth_state')
        service_account_file = data.get('service_account_file')
        service_account_json = data.get('service_account_json')

        print(f"🔍 performed_by_name value: '{performed_by_name}'")  # Debug logging

        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        if not performed_by_name:
            return jsonify({'success': False, 'error': 'Performed By Name is required'}), 400

        # Authenticate
        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400

        # Connect to sheet and get IDs from column
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        columns = sheets_service.get_columns(worksheet)

        if column_index >= len(columns):
            return jsonify({'success': False, 'error': f'Column index {column_index} not found'}), 400

        # Get IDs from column (skip header row)
        ids = [id.strip() for id in columns[column_index][1:] if id.strip()]

        if not ids:
            return jsonify({'success': False, 'error': 'No IDs found in the specified column'}), 400

        results = automation_service.add_privileges(ids, app_name, comment, performed_by_name)
        
        return jsonify({
            'success': True,
            'results': results,
            'total': len(results),
            'successful': sum(1 for r in results if r.get('success'))
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/revoke', methods=['POST'])
def revoke_privileges():
    """Revoke privileges from users - reads from Google Sheets column F (index 5)"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service
        automation_service = user_session.automation_service

        data = request.json
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        app_name = data.get('app_name', 'SLATE')
        comment = data.get('comment', '')
        column_index = data.get('column_index', 5)  # Column F (index 5)
        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        # Authenticate
        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400

        # Connect to sheet and get IDs from column
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        columns = sheets_service.get_columns(worksheet)

        if column_index >= len(columns):
            return jsonify({'success': False, 'error': f'Column index {column_index} not found'}), 400

        # Get IDs from column (skip header row)
        ids = [id.strip() for id in columns[column_index][1:] if id.strip()]

        if not ids:
            return jsonify({'success': False, 'error': 'No IDs found in the specified column'}), 400

        results = automation_service.revoke_privileges(ids, app_name, comment)
        
        return jsonify({
            'success': True,
            'results': results,
            'total': len(results),
            'successful': sum(1 for r in results if r.get('success'))
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

STATUS_FIELD_TO_RESULT = {
    'SOURCE_SYSTEM': 'source',
    'EMPLOYMENT_STATUS': 'employment_status',
    'STUDENT_STATUS_CODE': 'student_status_code',
}
STATUS_FIELD_HEADERS = {
    'SOURCE_SYSTEM': 'Source System',
    'EMPLOYMENT_STATUS': 'Employment Status',
    'STUDENT_STATUS_CODE': 'Student Status Code',
}

@app.route('/api/automation/get-employment-status', methods=['POST'])
def get_employment_status():
    """Get status fields (Source System, Employment Status, Student Status Code) - writes to Google Sheets"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service
        automation_service = user_session.automation_service

        data = request.json
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        ids = data.get('ids', [])
        id_type = data.get('id_type', 'SID')
        column_index = data.get('column_index', 0)
        to_fields = data.get('to_fields', ['SOURCE_SYSTEM', 'EMPLOYMENT_STATUS'])
        write_columns = data.get('write_columns', {})
        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400
        if not to_fields or not write_columns:
            return jsonify({'success': False, 'error': 'Select at least one field to get and a column for each'}), 400
        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        if not ids:
            columns = sheets_service.get_columns(worksheet)
            if column_index >= len(columns):
                return jsonify({'success': False, 'error': f'Column index {column_index} not found'}), 400
            ids = [id.strip() for id in columns[column_index][1:] if id.strip()]
        if not ids:
            return jsonify({'success': False, 'error': 'No IDs provided or found'}), 400
        print(f"📊 Starting status check for {len(ids)} users, fields: {to_fields}")
        for f, col in write_columns.items():
            if f not in to_fields:
                continue
            title = STATUS_FIELD_HEADERS.get(f, f)
            try:
                worksheet.update(f'{col}1', [[title]])
                print(f"  📌 Header {col}1 = '{title}'")
            except Exception as e:
                print(f"  ⚠️  Failed to write header {col}1: {str(e)}")
        def update_sheet_callback(index, result):
            row_number = index + 2
            fill = 'Not found' if not result.get('success') else None
            for f, col in write_columns.items():
                if f not in to_fields:
                    continue
                key = STATUS_FIELD_TO_RESULT.get(f, f.lower())
                val = fill if fill else result.get(key, '') or ''
                if fill is None and result.get('success'):
                    print(f"  ✅ [{index+1}/{len(ids)}] {result.get('id')} {f}='{val}'")
                try:
                    worksheet.update(f'{col}{row_number}', [[val]])
                except Exception as e:
                    print(f"  ⚠️  Failed to update {col}{row_number}: {str(e)}")
        results = automation_service.get_employment_status(
            ids, id_type, to_fields=to_fields, on_result_callback=update_sheet_callback
        )
        print(f"✅ Completed status check for {len(ids)} users")
        return jsonify({
            'success': True,
            'results': results,
            'message': f'Processed {len(ids)} users with real-time sheet updates'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/convert-id', methods=['POST'])
def convert_id():
    """Convert between ID types (Net ID to SID, Brown ID to SID)"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service
        automation_service = user_session.automation_service

        data = request.json
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        ids = data.get('ids', [])
        from_type = data.get('from_type', 'NETID')
        to_types = data.get('to_types', ['SID'])
        write_columns = data.get('write_columns', {'SID': 'B'})
        column_index = data.get('column_index', 0)

        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        # Authenticate
        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400
        
        # Connect to sheet
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        
        # If IDs not provided, read from sheet
        if not ids:
            columns = sheets_service.get_columns(worksheet)
            if column_index >= len(columns):
                return jsonify({'success': False, 'error': f'Column index {column_index} not found'}), 400
            ids = [id.strip() for id in columns[column_index][1:] if id.strip()]
        
        if not ids:
            return jsonify({'success': False, 'error': 'No IDs provided or found'}), 400
        
        if not to_types or not write_columns:
            return jsonify({'success': False, 'error': 'At least one target type and column required'}), 400
        print(f"📊 Starting ID conversion for {len(ids)} users ({from_type} → {', '.join(to_types)})")
        
        # Row 1 headers only for write columns (do not change the read column)
        CONVERT_TYPE_HEADERS = {
            'SID': 'Short ID',
            'NETID': 'Net ID',
            'BID': 'Brown ID',
            'BROWN_EMAIL': 'Brown Email',
        }
        for t, col in write_columns.items():
            if t not in to_types:
                continue
            title = CONVERT_TYPE_HEADERS.get(t, t)
            try:
                worksheet.update(f'{col}1', [[title]])
                print(f"  📌 Header {col}1 = '{title}'")
            except Exception as e:
                print(f"  ⚠️  Failed to write header {col}1: {str(e)}")
        
        # Define callback to update sheet after each result
        def update_sheet_callback(index, result):
            row_number = index + 2
            if result.get('success'):
                converted_ids = result.get('converted_ids', {})
                print(f"  ✅ [{index+1}/{len(ids)}] {result.get('id')} → {converted_ids}")
                for t, col in write_columns.items():
                    if t not in to_types:
                        continue
                    val = converted_ids.get(t, '')
                    try:
                        cell = f'{col}{row_number}'
                        worksheet.update(cell, [[val]])
                        print(f"  📝 {cell}='{val}'")
                    except Exception as e:
                        print(f"  ⚠️  Failed to update {cell}: {str(e)}")
            else:
                fill_val = 'Not found'
                print(f"  ❌ [{index+1}/{len(ids)}] {result.get('id')}: {result.get('error', 'Unknown error')}")
                for t, col in write_columns.items():
                    if t not in to_types:
                        continue
                    try:
                        worksheet.update(f'{col}{row_number}', [[fill_val]])
                    except Exception as e:
                        print(f"  ⚠️  Failed to update sheet for row {row_number}: {str(e)}")
        
        results = automation_service.convert_ids(ids, from_type, to_types, write_columns, on_result_callback=update_sheet_callback)
        
        print(f"✅ Completed ID conversion for {len(ids)} users")
        
        return jsonify({
            'success': True,
            'results': results,
            'message': f'Converted {len(ids)} IDs with real-time sheet updates'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/convert-validation', methods=['POST'])
def convert_validation():
    """Validate MyAccount search results against a selected sheet column and color cells."""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service
        automation_service = user_session.automation_service

        data = request.json or {}
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        search_mappings = data.get('search_mappings', [])
        check_column = str(data.get('check_column', 'X')).strip().upper()
        data_start_row = int(data.get('data_start_row', 2))

        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        if not isinstance(search_mappings, list):
            return jsonify({'success': False, 'error': 'search_mappings must be a list'}), 400

        normalized_mappings = []
        for item in search_mappings:
            if not isinstance(item, dict):
                continue
            field = str(item.get('search_field', '')).strip().upper()
            col = str(item.get('column', '')).strip().upper()
            if not field or not col:
                continue
            normalized_mappings.append({'search_field': field, 'column': col})

        if not normalized_mappings:
            return jsonify({'success': False, 'error': 'At least one search mapping is required'}), 400

        if not check_column:
            return jsonify({'success': False, 'error': 'Check column is required'}), 400

        if data_start_row < 1:
            return jsonify({'success': False, 'error': 'data_start_row must be >= 1'}), 400

        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400

        worksheet = sheets_service.connect(sheet_url, sheet_name)
        columns = sheets_service.get_columns(worksheet)

        def col_letter_to_index(letter):
            idx = 0
            for ch in letter:
                if ch < 'A' or ch > 'Z':
                    raise ValueError(f'Invalid column letter: {letter}')
                idx = idx * 26 + (ord(ch) - ord('A') + 1)
            return idx - 1

        def _clean(v):
            text = str(v or '')
            text = text.replace('\u00a0', ' ').replace('\u200b', '')
            text = ' '.join(text.split()).strip().lower()
            return text

        def comparison_keys(v):
            base = _clean(v)
            if not base:
                return set()

            keys = {base, base.replace(' ', '')}

            alnum_only = re.sub(r'[^a-z0-9]', '', base)
            if alnum_only:
                keys.add(alnum_only)

            if '@' in base:
                local_part = base.split('@')[0].strip()
                if local_part:
                    keys.add(local_part)
                    keys.add(re.sub(r'[^a-z0-9]', '', local_part))

            if ':' in base:
                after_colon = base.split(':', 1)[1].strip()
                if after_colon:
                    keys.add(after_colon)
                    keys.add(after_colon.replace(' ', ''))
                    keys.add(re.sub(r'[^a-z0-9]', '', after_colon))

            numeric = base.replace(',', '')
            if re.fullmatch(r'\d+\.0+', numeric):
                keys.add(numeric.split('.', 1)[0])

            compact = base.replace(' ', '').replace('_', '').replace('-', '')
            if compact:
                keys.add(compact)

            return {k for k in keys if k}

        try:
            mapping_indices = [
                {
                    'search_field': m['search_field'],
                    'column': m['column'],
                    'index': col_letter_to_index(m['column'])
                }
                for m in normalized_mappings
            ]
            check_index = col_letter_to_index(check_column)
        except ValueError as e:
            return jsonify({'success': False, 'error': str(e)}), 400

        for m in mapping_indices:
            if m['index'] >= len(columns):
                return jsonify({'success': False, 'error': f"Source column {m['column']} not found in worksheet"}), 400
        if check_index >= len(columns):
            return jsonify({'success': False, 'error': f'Check column {check_column} not found in worksheet'}), 400

        lookup_items = []
        max_rows = max((len(columns[m['index']]) for m in mapping_indices), default=0)

        for row_number in range(data_start_row, max_rows + 1):
            search_values = {}
            source_cells = []
            for m in mapping_indices:
                col_values = columns[m['index']]
                cell_value = str(col_values[row_number - 1]).strip() if (row_number - 1) < len(col_values) else ''
                if cell_value:
                    search_values[m['search_field']] = cell_value
                    source_cells.append(f"{m['column']}{row_number}")

            if not search_values:
                continue

            lookup_items.append({
                'row': row_number,
                'search_values': search_values,
                'source_cells': source_cells
            })

        if not lookup_items:
            return jsonify({'success': False, 'error': 'No values found in selected source columns'}), 400

        check_values = columns[check_index]

        check_key_to_cells = {}
        check_cell_to_value = {}
        for row_number in range(data_start_row, len(check_values) + 1):
            idx = row_number - 1
            value = str(check_values[idx]).strip() if idx < len(check_values) else ''
            if not value:
                continue
            cell = f'{check_column}{row_number}'
            check_cell_to_value[cell] = value
            for key in comparison_keys(value):
                if key not in check_key_to_cells:
                    check_key_to_cells[key] = set()
                check_key_to_cells[key].add(cell)

        green_source_cells = set()
        red_source_cells = set()
        green_check_cells = set()

        def process_validation_result(item_result):
            source_cells = item_result.get('source_cells', []) or []
            extracted_values = item_result.get('extracted_values', []) or []
            matched_check_cells = set()
            matched_values = []

            for extracted in extracted_values:
                extracted_keys = comparison_keys(extracted)
                extracted_match_cells = set()
                for key in extracted_keys:
                    extracted_match_cells.update(check_key_to_cells.get(key, set()))
                if extracted_match_cells:
                    matched_values.append(extracted)
                    matched_check_cells.update(extracted_match_cells)

            item_result['match_scope'] = 'column-wide'
            item_result['row_check_value'] = ''
            item_result['row_check_keys'] = []
            item_result['extracted_keys'] = {
                str(extracted): sorted(list(comparison_keys(extracted)))
                for extracted in extracted_values
            }
            item_result['matched_check_values'] = [
                check_cell_to_value[cell]
                for cell in sorted(list(matched_check_cells))
                if cell in check_cell_to_value
            ]

            item_result['matched_values'] = matched_values
            item_result['matched_check_cells'] = sorted(list(matched_check_cells))

            if source_cells:
                if item_result.get('success') and matched_check_cells:
                    green_source_cells.update(source_cells)
                    green_check_cells.update(matched_check_cells)

                    # Realtime green updates per iteration
                    try:
                        sheets_service.color_cells_green(worksheet, sorted(list(set(source_cells))))
                        sheets_service.color_cells_green(worksheet, sorted(list(matched_check_cells)))
                    except Exception as e:
                        print(f"⚠️  Realtime green color update failed: {e}")
                else:
                    red_source_cells.update(source_cells)

                    # Realtime red updates per iteration
                    try:
                        sheets_service.color_cells_red(worksheet, sorted(list(set(source_cells))))
                    except Exception as e:
                        print(f"⚠️  Realtime red color update failed: {e}")

        def on_validation_result(index, result):
            try:
                process_validation_result(result)
            except Exception as e:
                print(f"⚠️  Callback processing failed for index {index}: {e}")

        validation_results = automation_service.run_conversion_validation(
            lookup_items,
            on_result_callback=on_validation_result
        )

        # Ensure all results are processed even if callback missed any.
        for item_result in validation_results:
            if 'matched_check_cells' not in item_result:
                process_validation_result(item_result)

        red_source_cells = red_source_cells - green_source_cells

        matched_count = sum(1 for r in validation_results if r.get('matched_check_cells'))

        return jsonify({
            'success': True,
            'results': validation_results,
            'processed': len(validation_results),
            'matched': matched_count,
            'unmatched': len(validation_results) - matched_count,
            'green_source_cells': len(green_source_cells),
            'green_check_cells': len(green_check_cells),
            'red_source_cells': len(red_source_cells),
            'data_start_row': data_start_row,
            'message': f'Validation complete: {matched_count} matched, {len(validation_results) - matched_count} unmatched.'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/compare-lists', methods=['POST'])
def compare_lists():
    """Compare two lists to find who to add and remove"""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service

        data = request.json
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        list1 = data.get('list1', [])
        list2 = data.get('list2', [])
        column1_index = data.get('column1_index', 0)  # Column A (index 0)
        column2_index = data.get('column2_index', 1)  # Column B (index 1)
        to_add_column = data.get('to_add_column', 'E')  # Column for "To Add" results
        to_remove_column = data.get('to_remove_column', 'F')  # Column for "To Remove" results

        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        # Authenticate
        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400
        
        # Connect to sheet
        worksheet = sheets_service.connect(sheet_url, sheet_name)
        columns = sheets_service.get_columns(worksheet)
        
        # If lists not provided, read from sheet
        if not list1 and column1_index < len(columns):
            list1 = [item.strip() for item in columns[column1_index][1:] if item.strip()]
        if not list2 and column2_index < len(columns):
            list2 = [item.strip() for item in columns[column2_index][1:] if item.strip()]
        
        # Create lowercase versions for case-insensitive comparison
        list1_lower = [item.lower() for item in list1]
        list2_lower = [item.lower() for item in list2]
        
        # Case-insensitive comparison (keep original casing in output)
        to_add = [item for item in list1 if item.lower() not in list2_lower and item != ""]
        to_remove = [item for item in list2 if item.lower() not in list1_lower and item != ""]
        
        # Write results to specified columns using batch update
        updates = []
        
        # To Add column
        updates.append({'range': f'{to_add_column}1', 'values': [['To Add']]})
        if to_add:
            add_range = f'{to_add_column}2:{to_add_column}{len(to_add) + 1}'
            add_values = [[item] for item in to_add]
            updates.append({'range': add_range, 'values': add_values})
        
        # To Remove column
        updates.append({'range': f'{to_remove_column}1', 'values': [['To Remove']]})
        if to_remove:
            remove_range = f'{to_remove_column}2:{to_remove_column}{len(to_remove) + 1}'
            remove_values = [[item] for item in to_remove]
            updates.append({'range': remove_range, 'values': remove_values})
        
        # Execute batch update
        if updates:
            sheets_service.update_cells_batch(worksheet, updates)
        
        return jsonify({
            'success': True,
            'to_add': to_add,
            'to_remove': to_remove,
            'to_add_count': len(to_add),
            'to_remove_count': len(to_remove),
            'message': f'Results written to columns {to_add_column} and {to_remove_column}'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/move-and-shift-columns', methods=['POST'])
def move_and_shift_columns():
    """Move selected column values to destination columns when conditions match, then compact source columns upward."""
    try:
        user_session = _get_user_session()
        sheets_service = user_session.sheets_service

        data = request.json or {}
        sheet_url = data.get('sheet_url')
        sheet_name = data.get('sheet_name')
        source_columns = data.get('source_columns', [])
        destination_columns = data.get('destination_columns', [])
        conditions = data.get('conditions', [])
        destination_start_row = int(data.get('destination_start_row', 2))
        data_start_row = int(data.get('data_start_row', 2))

        if not sheet_url or not sheet_name:
            return jsonify({'success': False, 'error': 'Sheet URL and sheet name are required'}), 400

        if not source_columns or not destination_columns:
            return jsonify({'success': False, 'error': 'Source and destination columns are required'}), 400

        if len(source_columns) != len(destination_columns):
            return jsonify({'success': False, 'error': 'Source and destination columns must have the same count'}), 400

        if destination_start_row < 1 or data_start_row < 1:
            return jsonify({'success': False, 'error': 'Row numbers must be >= 1'}), 400

        normalized_source_columns = [str(c).strip().upper() for c in source_columns if str(c).strip()]
        normalized_destination_columns = [str(c).strip().upper() for c in destination_columns if str(c).strip()]

        if len(normalized_source_columns) != len(source_columns) or len(normalized_destination_columns) != len(destination_columns):
            return jsonify({'success': False, 'error': 'Columns cannot be empty'}), 400

        def column_letter_to_index(letter):
            idx = 0
            for ch in letter:
                if not ('A' <= ch <= 'Z'):
                    raise ValueError(f'Invalid column letter: {letter}')
                idx = idx * 26 + (ord(ch) - ord('A') + 1)
            return idx

        supported_operators = {
            'equals',
            'not_equals',
            'contains',
            'not_contains',
            'starts_with',
            'ends_with',
            'is_empty',
            'is_not_empty',
        }

        normalized_conditions = []
        for cond in conditions:
            col = str(cond.get('column', '')).strip().upper()
            op = str(cond.get('operator', 'equals')).strip().lower()
            val = str(cond.get('value', '')).strip()

            if not col:
                continue

            if op not in supported_operators:
                return jsonify({'success': False, 'error': f'Unsupported operator: {op}'}), 400

            if op not in {'is_empty', 'is_not_empty'} and val == '':
                return jsonify({'success': False, 'error': f'Condition value is required for operator {op} (column {col})'}), 400

            normalized_conditions.append({'column': col, 'operator': op, 'value': val})

        if not normalized_conditions:
            return jsonify({'success': False, 'error': 'At least one condition is required'}), 400

        try:
            authenticate_sheets_service(sheets_service, data)
        except Exception as e:
            return jsonify({'success': False, 'error': str(e)}), 400

        worksheet = sheets_service.connect(sheet_url, sheet_name)
        rows = worksheet.get_all_values()

        source_idx = [column_letter_to_index(c) - 1 for c in normalized_source_columns]
        condition_idx = [column_letter_to_index(c['column']) - 1 for c in normalized_conditions]

        last_relevant_row = data_start_row - 1
        for row_number, row in enumerate(rows, start=1):
            if row_number < data_start_row:
                continue
            relevant_columns = source_idx + condition_idx
            for idx in relevant_columns:
                if idx < len(row) and str(row[idx]).strip() != '':
                    last_relevant_row = row_number
                    break

        if last_relevant_row < data_start_row:
            return jsonify({
                'success': True,
                'moved_count': 0,
                'message': 'No source data found in selected columns.'
            })

        moved_records = []
        remaining_records = []

        for row_number in range(data_start_row, last_relevant_row + 1):
            row = rows[row_number - 1] if row_number - 1 < len(rows) else []

            src_values = []
            for idx in source_idx:
                src_values.append(str(row[idx]).strip() if idx < len(row) else '')

            has_source_data = any(v != '' for v in src_values)

            def matches_condition(current_val, condition):
                op = condition['operator']
                expected = condition['value']

                if op == 'equals':
                    return current_val == expected
                if op == 'not_equals':
                    return current_val != expected
                if op == 'contains':
                    return expected in current_val
                if op == 'not_contains':
                    return expected not in current_val
                if op == 'starts_with':
                    return current_val.startswith(expected)
                if op == 'ends_with':
                    return current_val.endswith(expected)
                if op == 'is_empty':
                    return current_val == ''
                if op == 'is_not_empty':
                    return current_val != ''
                return False

            condition_match = True
            for cond, idx in zip(normalized_conditions, condition_idx):
                current_val = str(row[idx]).strip() if idx < len(row) else ''
                if not matches_condition(current_val, cond):
                    condition_match = False
                    break

            if condition_match and has_source_data:
                moved_records.append(src_values)
            elif has_source_data:
                remaining_records.append(src_values)

        updates = []

        if moved_records:
            for src_col_pos, dest_col in enumerate(normalized_destination_columns):
                dest_values = [[record[src_col_pos]] for record in moved_records]
                start = destination_start_row
                end = destination_start_row + len(dest_values) - 1
                updates.append({
                    'range': f'{dest_col}{start}:{dest_col}{end}',
                    'values': dest_values
                })

        source_span = last_relevant_row - data_start_row + 1
        for src_col_pos, src_col in enumerate(normalized_source_columns):
            compacted_values = [[record[src_col_pos]] for record in remaining_records]
            while len(compacted_values) < source_span:
                compacted_values.append([''])

            updates.append({
                'range': f'{src_col}{data_start_row}:{src_col}{last_relevant_row}',
                'values': compacted_values
            })

        if updates:
            sheets_service.update_cells_batch(worksheet, updates)

        return jsonify({
            'success': True,
            'moved_count': len(moved_records),
            'remaining_count': len(remaining_records),
            'processed_rows': source_span,
            'message': f"Moved {len(moved_records)} row(s) and compacted source columns {', '.join(normalized_source_columns)}"
        })
    except ValueError as e:
        return jsonify({'success': False, 'error': str(e)}), 400
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/automation/logout', methods=['POST'])
def logout():
    """Logout and close browser, and drop this user's session."""
    try:
        sid = session.get('sid')
        if sid:
            session_manager.remove(sid)
            session.pop('sid', None)
        return jsonify({'success': True, 'message': 'Logged out successfully'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/')
def serve_index():
    """Serve frontend so the app can be used at http://localhost:5001/"""
    if os.path.isdir(FRONTEND_DIR):
        return send_from_directory(FRONTEND_DIR, 'index.html')
    return jsonify({'error': 'Frontend not found'}), 404

@app.route('/<path:path>')
def serve_frontend(path):
    """Serve frontend static files - registered last so API routes take precedence."""
    if os.path.isdir(FRONTEND_DIR) and path and not path.startswith('api'):
        return send_from_directory(FRONTEND_DIR, path)
    return jsonify({'error': 'Not found'}), 404

if __name__ == '__main__':
    app.run(debug=config.FLASK_DEBUG, port=config.PORT, host='0.0.0.0')
