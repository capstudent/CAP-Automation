from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time
import platform


class OperationAbortedException(Exception):
    """Raised when the user aborts an operation mid-run"""
    pass


class AutomationService:
    def __init__(self, config, sheets_service):
        self.config = config
        self.sheets_service = sheets_service
        self.driver = None
        self.wait = None
        self.abort_requested = False
    
    def abort(self):
        """Signal to abort the current operation. Navigator will return to myaccount."""
        self.abort_requested = True
    
    def _check_abort(self):
        """If abort was requested, navigate back to myaccount and raise."""
        if self.abort_requested:
            self.abort_requested = False
            self._navigate_to_myaccount()
            raise OperationAbortedException("Operation aborted by user")
    
    def _navigate_to_myaccount(self):
        """Navigate the browser back to the default myaccount person/search page."""
        if self.driver and self._is_driver_alive():
            try:
                self.driver.get("https://myaccount.brown.edu/person/search")
                print("↩️  Navigated back to myaccount")
            except Exception as e:
                print(f"⚠️  Could not navigate to myaccount: {str(e)[:100]}")
    
    def _is_driver_alive(self):
        """Check if the driver is still alive and responsive"""
        if not self.driver:
            return False
        try:
            # Try to get current URL - if this fails, driver is dead
            _ = self.driver.current_url
            return True
        except Exception:
            return False
    
    def _cleanup_driver(self):
        """Clean up the driver if it exists"""
        if self.driver:
            try:
                self.driver.quit()
            except Exception as e:
                print(f"Note: Error while closing driver: {str(e)[:100]}")
            finally:
                self.driver = None
                self.wait = None
    
    def _setup_driver(self):
        """Setup Chrome driver with appropriate options"""
        # If driver exists but is not responsive, clean it up
        if self.driver and not self._is_driver_alive():
            print(f"Detected unresponsive driver, cleaning up...")
            self._cleanup_driver()
        
        if self.driver:
            return self.driver
        
        chrome_options = Options()
        
        # Headless mode for server environments
        # Set to False if you want to see the browser (useful for debugging)
        # TEMPORARILY DISABLED for Duo debugging
        # chrome_options.add_argument('--headless=new')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-software-rasterizer')
        
        # Use a random port for remote debugging to avoid port conflicts
        import random
        debug_port = random.randint(9000, 9999)
        chrome_options.add_argument(f'--remote-debugging-port={debug_port}')
        
        # Additional options for stability and to avoid Chrome user data conflicts
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-popup-blocking')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # DISABLE ALL CREDENTIAL/PASSWORD/PASSKEY/WEBAUTHN POPUPS - AGGRESSIVE MODE
        # Disable multiple features in one argument (more effective)
        chrome_options.add_argument('--disable-features=PasswordManager,Credentials,WebAuthn,WebAuthenticationUI,WebAuthenticationRemoteDesktopSupport')
        
        # Additional WebAuthn disabling arguments
        chrome_options.add_argument('--disable-web-security')  # Helps disable WebAuthn
        chrome_options.add_argument('--disable-features=VirtualAuthenticatorAPI')
        
        # Set preferences to disable password and credential prompts
        prefs = {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "profile.default_content_setting_values.notifications": 2,  # Block notifications
            "autofill.profile_enabled": False,
            "autofill.credit_card_enabled": False,
            # Disable payment methods and autofill
            "payments.can_make_payment_enabled": False,
            # Explicitly disable WebAuthn/passkey prompts
            "webauthn.enable_credential_management": False,
            # Disable save password prompts
            "credentials_enable_autosignin": False,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Use a temporary user data directory to avoid conflicts
        import tempfile
        user_data_dir = tempfile.mkdtemp(prefix='chrome_automation_')
        chrome_options.add_argument(f'--user-data-dir={user_data_dir}')
        
        print(f"Starting Chrome with debugging port {debug_port}...")
        
        # Use webdriver-manager to automatically handle ChromeDriver
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            print(f"✓ Chrome driver initialized successfully")
        except Exception as e:
            # Fallback: try system chromedriver
            print(f"WebDriver manager failed, trying system chromedriver...")
            try:
                self.driver = webdriver.Chrome(options=chrome_options)
            except Exception as e2:
                raise Exception(f"Failed to initialize Chrome driver: {e2}. Make sure Chrome and ChromeDriver are installed.")
        
        # CRITICAL: Override WebAuthn/Credential Management APIs GLOBALLY
        # This prevents the browser-native popup that can't be inspected
        print(f"🔒 Disabling WebAuthn APIs globally...")
        try:
            # Execute CDP command to add script that runs on every page load
            self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': '''
                    // Completely disable WebAuthn API
                    Object.defineProperty(navigator, 'credentials', {
                        get: () => undefined
                    });
                    
                    // Override PublicKeyCredential
                    if (window.PublicKeyCredential) {
                        window.PublicKeyCredential = undefined;
                    }
                    
                    // Disable credential management
                    if (navigator.credentials) {
                        navigator.credentials.get = () => Promise.reject(new Error('Credentials API disabled'));
                        navigator.credentials.create = () => Promise.reject(new Error('Credentials API disabled'));
                        navigator.credentials.store = () => Promise.reject(new Error('Credentials API disabled'));
                    }
                    
                    console.log('WebAuthn APIs disabled globally');
                '''
            })
            print(f"✅ WebAuthn APIs disabled via CDP")
        except Exception as e:
            print(f"⚠️  CDP override failed: {str(e)[:100]}")
        
        self.wait = WebDriverWait(self.driver, 20)
        return self.driver
    
    def _handle_duo_push_selection(self):
        """Helper to select Duo Push method and handle device trust"""
        try:
            print(f"\n{'='*60}")
            print(f"Attempting to select Duo Push method...")
            print(f"{'='*60}")
            
            # Wait a bit for Duo to fully load
            time.sleep(2)
            
            # Method 1: Try switching to iframe first (most common)
            try:
                print(f"Method 1: Looking for Duo iframe...")
                iframe = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, "duo_iframe"))
                )
                print(f"✓ Found Duo iframe, switching to it...")
                self.driver.switch_to.frame(iframe)
                
                # Wait for Duo content to load inside iframe
                print(f"⏱️  Waiting 3 seconds for Duo content to fully load...")
                time.sleep(3)
                
                # STEP 1: Look for popup/dialog/modal and click on it to focus it
                print(f"🎯 Looking for popup/dialog to focus...")
                popup_focused = False
                popup_selectors = [
                    (By.XPATH, "//*[contains(@class, 'dialog')]"),
                    (By.XPATH, "//*[contains(@class, 'modal')]"),
                    (By.XPATH, "//*[contains(@class, 'popup')]"),
                    (By.XPATH, "//*[contains(@role, 'dialog')]"),
                    (By.XPATH, "//*[contains(@class, 'passkey')]"),
                    (By.XPATH, "//*[contains(@class, 'webauthn')]"),
                ]
                
                for selector_type, selector_value in popup_selectors:
                    try:
                        popup = self.driver.find_element(selector_type, selector_value)
                        if popup.is_displayed():
                            print(f"✓ Found popup element, clicking to focus it...")
                            popup.click()
                            popup_focused = True
                            time.sleep(0.5)
                            break
                    except:
                        continue
                
                if popup_focused:
                    print(f"✓ Popup focused, ready for ESC key")
                else:
                    print(f"ℹ️  No popup found to focus (may not need it)")
                
                # STEP 2: Press ESC key 2-3 times to dismiss any random popups
                print(f"🔑 Pressing ESC key to dismiss any popups...")
                try:
                    body = self.driver.find_element(By.TAG_NAME, "body")
                    for i in range(3):
                        body.send_keys(Keys.ESCAPE)
                        time.sleep(0.5)
                    print(f"✓ Pressed ESC 3 times")
                except Exception as e:
                    print(f"⚠️  Could not press ESC key: {e}")
                
                time.sleep(1)
                
                # STEP 3: Look for and click any Close button in popups (if ESC didn't work)
                print(f"🔍 Looking for Close button in any popup...")
                close_selectors = [
                    (By.XPATH, "//button[contains(text(), 'Close')]"),
                    (By.XPATH, "//button[@aria-label='Close']"),
                    (By.CSS_SELECTOR, "button.close"),
                    (By.XPATH, "//*[contains(@class, 'close-button')]"),
                    (By.XPATH, "//button[contains(@class, 'close')]"),
                    (By.XPATH, "//*[@role='button'][contains(text(), 'Close')]"),
                ]
                
                popup_closed = False
                for selector_type, selector_value in close_selectors:
                    try:
                        close_btn = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((selector_type, selector_value))
                        )
                        close_btn.click()
                        print(f"✅ Clicked Close button in popup!")
                        popup_closed = True
                        time.sleep(1)
                        break
                    except:
                        continue
                
                if not popup_closed:
                    print(f"ℹ️  No Close button found (popup may have closed with ESC)")
                
                time.sleep(1)
                
                # Print what buttons are available for debugging
                try:
                    print(f"📋 Checking available buttons in Duo iframe...")
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    print(f"   Found {len(buttons)} buttons total")
                    for i, btn in enumerate(buttons):
                        btn_text = btn.text or btn.get_attribute("value") or btn.get_attribute("aria-label") or ""
                        btn_class = btn.get_attribute("class") or ""
                        btn_visible = btn.is_displayed()
                        if btn_text or btn_class:
                            print(f"   Button {i+1}: text='{btn_text}' class='{btn_class}' visible={btn_visible}")
                except Exception as e:
                    print(f"⚠️  Could not list buttons: {e}")
                
                # STEP 2: Try to find "Send Me a Push" button directly
                push_selectors = [
                    # Text-based selectors
                    (By.XPATH, "//button[contains(text(), 'Send Me a Push')]"),
                    (By.XPATH, "//button[contains(text(), 'Send me a push')]"),
                    (By.XPATH, "//button[contains(text(), 'Send Push')]"),
                    (By.XPATH, "//button[contains(., 'Push')]"),
                    (By.XPATH, "//button[contains(@value, 'push')]"),
                    
                    # Class-based selectors
                    (By.CSS_SELECTOR, "button.positive.auth-button"),
                    (By.CSS_SELECTOR, "button.auth-button[value='push']"),
                    (By.CSS_SELECTOR, "button[data-device-index='phone1']"),
                    
                    # Fallback: look for any button with "push" in class or value
                    (By.XPATH, "//button[contains(@class, 'push') or contains(@value, 'push')]"),
                ]
                
                push_found = False
                print(f"🔍 Trying {len(push_selectors)} different selectors for Push button...")
                for i, (selector_type, selector_value) in enumerate(push_selectors):
                    try:
                        print(f"   Selector {i+1}/{len(push_selectors)}: {selector_value[:60]}...")
                        push_button = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((selector_type, selector_value))
                        )
                        push_button.click()
                        print(f"✅ SUCCESS! Clicked Duo Push button directly")
                        print(f"   Used selector: {selector_value[:60]}")
                        push_found = True
                        break
                    except:
                        continue
                
                # STEP 4: If Push button not found, look for "Other options"
                if not push_found:
                    print(f"⚠️  Push button not immediately visible")
                    
                    # Try to find and click "Other options" link
                    print(f"🔄 Looking for 'Other options' link/button...")
                    
                    other_options_selectors = [
                        (By.XPATH, "//a[contains(text(), 'Other options')]"),
                        (By.XPATH, "//a[contains(text(), 'Other Options')]"),
                        (By.XPATH, "//button[contains(text(), 'Other options')]"),
                        (By.XPATH, "//button[contains(text(), 'Other Options')]"),
                        (By.CSS_SELECTOR, "a.other-options"),
                        (By.CSS_SELECTOR, "button.other-options"),
                        (By.XPATH, "//button[contains(@class, 'other-options')]"),
                        (By.XPATH, "//*[contains(text(), 'other options')]"),
                        (By.XPATH, "//*[contains(text(), 'Other options')]"),
                    ]
                    
                    other_options_clicked = False
                    for i, (selector_type, selector_value) in enumerate(other_options_selectors):
                        try:
                            print(f"   Trying {i+1}/{len(other_options_selectors)}: {selector_value[:60]}...")
                            other_options_btn = WebDriverWait(self.driver, 2).until(
                                EC.element_to_be_clickable((selector_type, selector_value))
                            )
                            other_options_btn.click()
                            print(f"✅ Clicked 'Other options' button")
                            other_options_clicked = True
                            time.sleep(2)  # Wait for options to expand
                            break
                        except:
                            continue
                    
                    if other_options_clicked:
                        # Now try to find "Duo Push" from the list
                        print(f"🔍 Looking for 'Duo Push' in expanded options...")
                        
                        # Check what's available now
                        try:
                            buttons = self.driver.find_elements(By.TAG_NAME, "button")
                            print(f"📋 Available options after clicking 'Other options':")
                            for i, btn in enumerate(buttons):
                                btn_text = btn.text or btn.get_attribute("value") or ""
                                btn_visible = btn.is_displayed()
                                if btn_text and btn_visible:
                                    print(f"   Option {i+1}: '{btn_text}'")
                        except:
                            pass
                        
                        # Try to find and click "Duo Push" option
                        duo_push_selectors = [
                            (By.XPATH, "//button[contains(., 'Duo Push')]"),
                            (By.XPATH, "//*[contains(text(), 'Duo Push')]"),
                            (By.XPATH, "//div[contains(text(), 'Duo Push')]"),
                            (By.XPATH, "//button[contains(., 'Send to')]"),  # "Send to USA"
                        ]
                        
                        for selector_type, selector_value in duo_push_selectors:
                            try:
                                duo_push_option = WebDriverWait(self.driver, 2).until(
                                    EC.element_to_be_clickable((selector_type, selector_value))
                                )
                                duo_push_option.click()
                                print(f"✅ SUCCESS! Clicked 'Duo Push' from expanded options")
                                push_found = True
                                break
                            except:
                                continue
                
                if not push_found:
                    print(f"❌ Could not find Duo Push button even after trying all methods")
                
                # Switch back to main content
                self.driver.switch_to.default_content()
                
                return push_found
                
            except Exception as e:
                print(f"Method 1 (iframe) failed: {str(e)[:150]}")
                # Make sure we're back to default content
                try:
                    self.driver.switch_to.default_content()
                except:
                    pass
            
            # Method 2: Try clicking push button directly (no iframe)
            try:
                print(f"\nMethod 2: Looking for push button in main content...")
                
                # Wait for page to fully load before dismissing popups
                print(f"⏱️  Waiting 3 seconds for Duo page to fully load...")
                time.sleep(3)
                
                # STEP 0: Use JavaScript to forcefully remove WebAuthn/passkey popups
                print(f"🔨 Using JavaScript to remove WebAuthn popups...")
                try:
                    # JavaScript to remove all WebAuthn-related elements
                    js_remove_popups = """
                    // Remove all elements containing webauthn, passkey, or credential classes
                    const selectors = [
                        '[class*="webauthn"]',
                        '[class*="passkey"]',
                        '[class*="credential"]',
                        '[role="dialog"]',
                        '[class*="dialog"]',
                        '[class*="modal"]'
                    ];
                    
                    selectors.forEach(selector => {
                        const elements = document.querySelectorAll(selector);
                        elements.forEach(el => {
                            // Only remove if it's related to authentication prompts
                            if (el.textContent.includes('passkey') || 
                                el.textContent.includes('Verify your identity') ||
                                el.className.includes('webauthn')) {
                                el.remove();
                                console.log('Removed popup element:', el.className);
                            }
                        });
                    });
                    
                    // Also cancel any WebAuthn credential requests
                    if (window.PublicKeyCredential) {
                        window.PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable = () => Promise.resolve(false);
                        window.PublicKeyCredential.isConditionalMediationAvailable = () => Promise.resolve(false);
                    }
                    """
                    self.driver.execute_script(js_remove_popups)
                    print(f"✓ JavaScript popup removal executed")
                    time.sleep(1)
                except Exception as e:
                    print(f"⚠️  JavaScript removal failed: {str(e)[:100]}")
                
                # STEP 1: Close "No passkeys available" dialog if present
                print(f"🔍 Looking for 'No passkeys available' dialog...")
                try:
                    # Look for the "Close" button in the passkey dialog
                    close_passkey_selectors = [
                        (By.XPATH, "//button[contains(text(), 'Close')]"),
                        (By.XPATH, "//button[text()='Close']"),
                        (By.CSS_SELECTOR, "button[class*='close']"),
                    ]
                    
                    passkey_closed = False
                    for selector_type, selector_value in close_passkey_selectors:
                        try:
                            close_btn = WebDriverWait(self.driver, 2).until(
                                EC.element_to_be_clickable((selector_type, selector_value))
                            )
                            print(f"✅ Found 'Close' button, clicking it...")
                            close_btn.click()
                            passkey_closed = True
                            time.sleep(1)
                            print(f"✓ Closed passkey dialog")
                            break
                        except:
                            continue
                    
                    if not passkey_closed:
                        print(f"ℹ️  No passkey dialog found")
                except Exception as e:
                    print(f"⚠️  Error closing passkey dialog: {str(e)[:100]}")
                
                # STEP 2: Press ESC multiple times to close Google Password Manager popup and any others
                print(f"🔑 Pressing ESC key 5 times to dismiss all popups...")
                try:
                    body = self.driver.find_element(By.TAG_NAME, "body")
                    for i in range(5):
                        body.send_keys(Keys.ESCAPE)
                        time.sleep(0.3)
                    print(f"✓ Pressed ESC 5 times")
                except Exception as e:
                    print(f"⚠️  Could not press ESC key: {e}")
                
                time.sleep(2)  # Wait for popups to close
                
                # STEP 3: Look for and click any remaining Close/X buttons
                print(f"🔍 Looking for any remaining Close/X buttons...")
                close_selectors = [
                    (By.XPATH, "//button[contains(text(), 'Close')]"),
                    (By.XPATH, "//button[@aria-label='Close']"),
                    (By.XPATH, "//button[contains(@aria-label, 'close')]"),
                    (By.CSS_SELECTOR, "button.close"),
                    (By.XPATH, "//*[contains(@class, 'close-button')]"),
                    (By.XPATH, "//button[contains(@class, 'close')]"),
                    (By.XPATH, "//*[@role='button'][contains(text(), 'Close')]"),
                    (By.XPATH, "//button[text()='×']"),  # X symbol
                    (By.XPATH, "//button[contains(text(), '×')]"),
                ]
                
                popup_closed = False
                for selector_type, selector_value in close_selectors:
                    try:
                        close_btn = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((selector_type, selector_value))
                        )
                        close_btn.click()
                        print(f"✅ Clicked remaining Close button!")
                        popup_closed = True
                        time.sleep(1)
                        break
                    except:
                        continue
                
                if not popup_closed:
                    print(f"✓ All popups closed (no more Close buttons found)")
                
                time.sleep(1)
                
                # IMPORTANT: Check if device trust dialog appeared BEFORE auth method selection
                # Sometimes Duo shows "Is this your device?" before letting you pick a method
                try:
                    print(f"🔍 Checking if device trust dialog appeared early...")
                    early_device_trust_btn = None
                    try:
                        early_device_trust_btn = self.driver.find_element(
                            By.XPATH, 
                            "//button[contains(text(), 'No, other people use this device')]"
                        )
                        if early_device_trust_btn.is_displayed():
                            print(f"✅ Found EARLY device trust dialog - clicking 'No, other people use this device'")
                            early_device_trust_btn.click()
                            print(f"✅ Clicked device trust button (early dialog)")
                            time.sleep(2)
                            # After clicking, the real auth method selection should appear
                    except:
                        print(f"ℹ️  No early device trust dialog")
                except Exception as e:
                    print(f"⚠️  Error checking early device trust: {str(e)[:100]}")
                
                # Print what buttons are available for debugging
                try:
                    print(f"📋 Checking available buttons in main content...")
                    buttons = self.driver.find_elements(By.TAG_NAME, "button")
                    print(f"   Found {len(buttons)} buttons total")
                    for i, btn in enumerate(buttons):
                        btn_text = btn.text or btn.get_attribute("value") or btn.get_attribute("aria-label") or ""
                        btn_class = btn.get_attribute("class") or ""
                        btn_visible = btn.is_displayed()
                        if btn_text or btn_class:
                            print(f"   Button {i+1}: text='{btn_text}' class='{btn_class}' visible={btn_visible}")
                except Exception as e:
                    print(f"⚠️  Could not list buttons: {e}")
                
                push_selectors = [
                    (By.XPATH, "//button[contains(text(), 'Send Me a Push')]"),
                    (By.XPATH, "//button[contains(., 'Send me a push')]"),
                    (By.CSS_SELECTOR, "button[value='push']"),
                    (By.XPATH, "//button[contains(@value, 'push')]"),
                ]
                
                push_found = False
                print(f"🔍 Trying {len(push_selectors)} different selectors for Push button...")
                for i, (selector_type, selector_value) in enumerate(push_selectors):
                    try:
                        print(f"   Selector {i+1}/{len(push_selectors)}: {selector_value[:60]}...")
                        push_button = WebDriverWait(self.driver, 1).until(
                            EC.element_to_be_clickable((selector_type, selector_value))
                        )
                        push_button.click()
                        print(f"✅ SUCCESS! Clicked Duo Push button (direct method)")
                        push_found = True
                        break
                    except:
                        continue
                
                # If Push button not found, look for "Other options"
                if not push_found:
                    print(f"⚠️  Push button not immediately visible")
                    print(f"🔄 Looking for 'Other options' link/button...")
                    
                    other_options_selectors = [
                        (By.XPATH, "//a[contains(text(), 'Other options')]"),
                        (By.XPATH, "//a[contains(text(), 'Other Options')]"),
                        (By.XPATH, "//button[contains(text(), 'Other options')]"),
                        (By.XPATH, "//button[contains(text(), 'Other Options')]"),
                        (By.CSS_SELECTOR, "a.other-options"),
                        (By.CSS_SELECTOR, "button.other-options"),
                        (By.XPATH, "//button[contains(@class, 'other-options')]"),
                        (By.XPATH, "//*[contains(text(), 'other options')]"),
                        (By.XPATH, "//*[contains(text(), 'Other options')]"),
                    ]
                    
                    other_options_clicked = False
                    for i, (selector_type, selector_value) in enumerate(other_options_selectors):
                        try:
                            print(f"   Trying {i+1}/{len(other_options_selectors)}: {selector_value[:60]}...")
                            other_options_btn = WebDriverWait(self.driver, 2).until(
                                EC.element_to_be_clickable((selector_type, selector_value))
                            )
                            other_options_btn.click()
                            print(f"✅ Clicked 'Other options' button")
                            other_options_clicked = True
                            print(f"⏱️  Waiting 1 second then removing WebAuthn popups...")
                            time.sleep(1)
                            
                            # CRITICAL: Remove WebAuthn popup that appears after clicking "Other options"
                            try:
                                js_remove_webauthn = """
                                // Remove WebAuthn popup elements that block the options menu
                                const webauthnElements = document.querySelectorAll('[class*="webauthn"], [class*="passkey"]');
                                webauthnElements.forEach(el => el.remove());
                                
                                // Also remove the "Verify your identity using this device" message
                                const messages = document.querySelectorAll('p.webauthn-request-message');
                                messages.forEach(el => el.parentElement.remove());
                                
                                console.log('Removed WebAuthn popup elements');
                                """
                                self.driver.execute_script(js_remove_webauthn)
                                print(f"✅ Removed WebAuthn popup with JavaScript")
                            except Exception as e:
                                print(f"⚠️  WebAuthn removal failed: {str(e)[:100]}")
                            
                            time.sleep(2)  # Wait for DOM to settle after removal
                            break
                        except:
                            continue
                    
                    if other_options_clicked:
                        # Now try to find "Duo Push" from the list
                        print(f"🔍 Looking for 'Duo Push' in expanded options...")
                        
                        # Check ALL elements, not just buttons
                        try:
                            # Check all buttons
                            buttons = self.driver.find_elements(By.TAG_NAME, "button")
                            print(f"📋 All buttons after clicking 'Other options':")
                            for i, btn in enumerate(buttons):
                                btn_text = btn.text or btn.get_attribute("value") or ""
                                btn_visible = btn.is_displayed()
                                if btn_text and btn_visible:
                                    print(f"   Button {i+1}: '{btn_text}'")
                            
                            # Also check for links and divs that might be clickable options
                            all_elements = self.driver.find_elements(By.XPATH, "//*[contains(@class, 'option') or contains(@class, 'method') or contains(@class, 'auth')]")
                            print(f"📋 All auth-related elements:")
                            for i, elem in enumerate(all_elements):
                                elem_text = elem.text or ""
                                elem_tag = elem.tag_name
                                elem_class = elem.get_attribute("class") or ""
                                elem_visible = elem.is_displayed()
                                if elem_visible and (elem_text or elem_class):
                                    print(f"   {elem_tag} {i+1}: text='{elem_text}' class='{elem_class}'")
                        except:
                            pass
                        
                        # Try to find and click "Duo Push" option
                        # IMPORTANT: Duo uses <a> links, not buttons!
                        duo_push_selectors = [
                            (By.XPATH, "//a[contains(@class, 'auth-method-link') and contains(., 'Duo Push')]"),
                            (By.XPATH, "//a[contains(text(), 'Duo Push')]"),
                            (By.XPATH, "//a[contains(., 'Duo Push')]"),
                            (By.XPATH, "//*[contains(@class, 'auth-method') and contains(., 'Duo Push')]"),
                            (By.XPATH, "//button[contains(text(), 'Duo Push')]"),
                            (By.XPATH, "//*[contains(text(), 'Duo Push')]"),
                        ]
                        
                        duo_push_clicked = False
                        for i, (selector_type, selector_value) in enumerate(duo_push_selectors):
                            try:
                                print(f"   Trying Duo Push selector {i+1}/{len(duo_push_selectors)}: {selector_value[:80]}...")
                                duo_push_option = WebDriverWait(self.driver, 2).until(
                                    EC.element_to_be_clickable((selector_type, selector_value))
                                )
                                duo_push_option.click()
                                print(f"✅ SUCCESS! Clicked 'Duo Push' using selector #{i+1}")
                                push_found = True
                                duo_push_clicked = True
                                break
                            except Exception as e:
                                continue
                        
                        if not duo_push_clicked:
                            print(f"❌ Failed to click 'Duo Push' - tried {len(duo_push_selectors)} selectors")
                
                if push_found:
                    return True
                        
            except Exception as e:
                print(f"Method 2 (direct) failed: {str(e)[:150]}")
            
            print(f"\n⚠️  WARNING: Could not find 'Send Me a Push' button")
            print(f"Duo may be using Touch ID/Fingerprint as default")
            print(f"You may need to manually select 'Duo Push' in your Duo settings")
            print(f"{'='*60}\n")
            return False
            
        except Exception as e:
            print(f"Error in Duo Push selection: {str(e)[:200]}")
            try:
                self.driver.switch_to.default_content()
            except:
                pass
            return False
    
    def login(self, username, password):
        """Login to myaccount.brown.edu"""
        if not self.driver:
            self._setup_driver()
        
        print(f"=== Starting login to myaccount.brown.edu ===")
        login_url = "https://myaccount.brown.edu/person/search"
        self.driver.get(login_url)
        print(f"Navigated to: {login_url}")
        
        # Wait for login form
        print(f"Waiting for login form...")
        try:
            username_field = self.wait.until(EC.presence_of_element_located((By.ID, "username")))
            password_field = self.wait.until(EC.presence_of_element_located((By.ID, "password")))
            submit_button = self.wait.until(EC.element_to_be_clickable((By.NAME, "_eventId_proceed")))
            print(f"Login form found")
            
            # Debug: Check form action
            try:
                form = self.driver.find_element(By.TAG_NAME, "form")
                form_action = form.get_attribute("action")
                form_method = form.get_attribute("method")
                print(f"Form action: {form_action}")
                print(f"Form method: {form_method}")
            except Exception as e:
                print(f"Could not get form details: {str(e)[:100]}")
                
        except Exception as e:
            print(f"ERROR: Could not find login form elements")
            print(f"Current URL: {self.driver.current_url}")
            print(f"Page title: {self.driver.title}")
            print(f"Exception: {str(e)}")
            return False
        
        # Enter credentials
        print(f"Entering credentials for user: {username}")
        username_field.clear()
        username_field.send_keys(username)
        time.sleep(0.5)
        password_field.clear()
        password_field.send_keys(password)
        time.sleep(0.5)
        
        # Store current URL before clicking submit
        url_before_submit = self.driver.current_url
        print(f"URL before submit: {url_before_submit}")
        
        # Try clicking the submit button
        try:
            # First, try clicking the button normally
            submit_button.click()
            print(f"Clicked submit button (normal click)")
        except Exception as e:
            print(f"Normal click failed: {str(e)[:100]}")
            # Try JavaScript click as alternative
            try:
                print(f"Trying JavaScript click...")
                self.driver.execute_script("arguments[0].click();", submit_button)
                print(f"Clicked submit button (JavaScript)")
            except Exception as e2:
                print(f"JavaScript click also failed: {str(e2)[:100]}")
                # Last resort: submit the form directly
                try:
                    print(f"Trying to submit form directly...")
                    form = self.driver.find_element(By.TAG_NAME, "form")
                    form.submit()
                    print(f"Submitted form directly")
                except Exception as e3:
                    print(f"Form submit also failed: {str(e3)[:100]}")
                    return False
        
        # Wait for page to change (URL should change after submit)
        try:
            print(f"Waiting for page navigation (up to 10 seconds)...")
            WebDriverWait(self.driver, 10).until(
                lambda driver: driver.current_url != url_before_submit
            )
            print(f"Page navigation detected")
        except Exception as e:
            print(f"Warning: URL did not change after submit")
            print(f"This might be normal if there's an error or if page uses AJAX")
        
        # Quick check for Duo immediately after navigation
        current_url = self.driver.current_url
        print(f"Current URL after submit: {current_url}")
        
        # Check if we got a blank/data page (something went wrong)
        if current_url.startswith("data:") or current_url == "about:blank":
            print(f"")
            print(f"=" * 60)
            print(f"ERROR: Got blank page after login submission!")
            print(f"=" * 60)
            print(f"Current URL: {current_url}")
            print(f"Page title: {self.driver.title}")
            print(f"")
            print(f"This usually means:")
            print(f"  1. The form submitted but navigated to empty page")
            print(f"  2. JavaScript might have prevented default form action")
            print(f"  3. The login page structure might have changed")
            print(f"")
            print(f"Possible solutions:")
            print(f"  - Check if myaccount.brown.edu is accessible")
            print(f"  - Verify credentials are correct")
            print(f"  - Check if page structure changed (view browser window)")
            print(f"")
            
            # Take a screenshot for debugging
            try:
                screenshot_path = "/tmp/login_error_screenshot.png"
                self.driver.save_screenshot(screenshot_path)
                print(f"Screenshot saved to: {screenshot_path}")
            except:
                pass
            
            return False
        
        # Check if we're on Duo page IMMEDIATELY
        if "duosecurity.com" in current_url:
            print(f"\n🔐 === Duo 2FA Detected === 🔐")
            print(f"Attempting to switch to Duo Push (instead of Touch ID)...")
            
            # Try to select Duo Push instead of fingerprint/other methods
            # Do this IMMEDIATELY before Duo auto-selects Touch ID
            self._handle_duo_push_selection()
            
            # Wait for Duo to redirect back (up to 60 seconds)
            print(f"Waiting for Duo approval (up to 60 seconds)...")
            print(f"👉 Check your Duo Mobile app and approve the login request!")
            
            # Enhanced wait: Check for BOTH Duo redirect AND device trust dialog
            duo_approved = False
            start_time = time.time()
            max_wait = 60
            
            while time.time() - start_time < max_wait:
                try:
                    current_url = self.driver.current_url
                    
                    # Check if we left Duo page
                    if "duosecurity.com" not in current_url:
                        print(f"✓ Duo authentication completed, redirected to: {current_url}")
                        duo_approved = True
                        break
                    
                    # Check for device trust dialog DURING the wait
                    # This appears after approving but before redirect
                    try:
                        device_trust_selectors = [
                            (By.XPATH, "//button[contains(text(), 'No, other people use this device')]"),
                            (By.XPATH, "//a[contains(text(), 'No, other people use this device')]"),
                            (By.XPATH, "//*[contains(text(), 'No, other people')]"),
                            (By.XPATH, "//button[contains(., 'other people')]"),
                            (By.ID, "dont-trust-browser-button"),
                        ]
                        
                        for selector_type, selector_value in device_trust_selectors:
                            try:
                                device_trust_btn = self.driver.find_element(selector_type, selector_value)
                                if device_trust_btn.is_displayed():
                                    print(f"🔍 Found device trust dialog!")
                                    device_trust_btn.click()
                                    print(f"✅ Clicked 'No, other people use this device' button")
                                    time.sleep(2)
                                    break
                            except:
                                continue
                    except:
                        pass
                    
                    time.sleep(0.5)  # Check every 500ms
                    
                except Exception as e:
                    # If we get a stale element or window closed error, might be transitioning
                    if "window" in str(e).lower() or "no such window" in str(e).lower():
                        print(f"⚠️  Window closed during Duo approval - this may be normal")
                        time.sleep(2)
                        break
                    time.sleep(0.5)
            
            if not duo_approved and time.time() - start_time >= max_wait:
                print(f"✗ Timeout waiting for Duo approval")
                return False
        
        # Wait for successful login - increase timeout for Duo approval
        # Try multiple ways to detect successful login
        try:
            # Wait up to 30 seconds for page to load after Duo
            wait_medium = WebDriverWait(self.driver, 30)
            
            print(f"Waiting for login success (up to 30 seconds)...")
            
            # Try multiple possible success indicators
            try:
                # Option 1: Look for "Search for a user" text
                print(f"Method 1: Looking for 'Search for a user' text...")
                wait_medium.until(EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Search for a user')]")
                ))
                print(f"Login successful! (Method 1)")
                return True
            except Exception as e1:
                print(f"Method 1 failed: {str(e1)[:100]}")
                # Option 2: Look for the search form
                try:
                    print(f"Method 2: Looking for search form...")
                    wait_medium.until(EC.presence_of_element_located((By.NAME, "search")))
                    print(f"Login successful! (Method 2)")
                    return True
                except Exception as e2:
                    print(f"Method 2 failed: {str(e2)[:100]}")
                    # Option 3: Check if URL changed to the person search page
                    print(f"Method 3: Checking URL...")
                    print(f"Current URL: {self.driver.current_url}")
                    if "person/search" in self.driver.current_url and "login" not in self.driver.current_url:
                        print(f"Login successful! (Method 3 - URL check)")
                        return True
                    print(f"Method 3 failed - URL doesn't indicate success")
                    
                    # Debug: Print page source snippet
                    print(f"Page title: {self.driver.title}")
                    print(f"Login failed - could not detect successful login")
                    return False
        except Exception as e:
            print(f"Login failed with exception: {str(e)}")
            print(f"Final URL: {self.driver.current_url}")
            return False
    
    def _get_next_attn_date(self):
        """Get the next 01/31 or 06/30 (whichever comes first in the future)"""
        now = datetime.now()
        year = now.year
        jan_31 = datetime(year, 1, 31)
        jun_30 = datetime(year, 6, 30)
        if now < jan_31:
            return f"01/31/{year}"
        elif now < jun_30:
            return f"06/30/{year}"
        else:
            return f"01/31/{year + 1}"
    
    def add_privileges(self, ids, app_name, comment, performed_by_name):
        """
        Add privileges to a list of user IDs
        
        Args:
            ids: List of user IDs
            app_name: Application name to grant privileges for
            comment: Comment text to add to the privilege record (displayed exactly as entered)
            performed_by_name: Name to enter in "Performed By" autocomplete field
        """
        if not self.driver:
            raise Exception("Not logged in. Please login first.")
        
        results = []
        
        try:
            for user_id in ids:
                self._check_abort()
                result = {'id': user_id, 'success': False, 'error': None}
                try:
                    # Search for user
                    search_button = self.driver.find_element(By.NAME, "search")
                    text_box = self.driver.find_element(By.NAME, "brown_login")
                    text_box.clear()
                    text_box.send_keys(user_id)
                    time.sleep(2)
                    search_button.click()
                    time.sleep(3)
                    
                    # Click View Overview
                    try:
                        vo = self.driver.find_element(
                            By.XPATH,
                            "//a[@class='btn btn-default' and contains(text(), 'View Overview')]"
                        )
                        vo.click()
                    except:
                        result['error'] = "Could not find View Overview"
                        results.append(result)
                        continue
                    
                    time.sleep(2)
                    
                    # Extract Employment Source (for Banner-specific logic)
                    employment_source = ""
                    try:
                        source_selectors = [
                            (By.XPATH, "//label[@class='col-xs-6' and contains(text(), 'Source')]/following-sibling::div[@class='col-xs-6']"),
                            (By.XPATH, "//b[@class='col-xs-6' and contains(text(), 'Source')]/following-sibling::div[@class='col-xs-6']"),
                            (By.XPATH, "//*[contains(text(), 'Source')]/following-sibling::div[@class='col-xs-6']"),
                        ]
                        for sel_type, sel_val in source_selectors:
                            try:
                                source_div = self.driver.find_element(sel_type, sel_val)
                                employment_source = (source_div.text or "").strip()
                                print(f"  📋 Employment Source: '{employment_source}'")
                                break
                            except:
                                continue
                    except Exception as e:
                        print(f"  ⚠️  Could not extract Employment Source: {str(e)[:80]}")
                    
                    # Navigate to AdminID - Current
                    button = self.driver.find_element(By.LINK_TEXT, "AdminID - Current")
                    button.click()
                    
                    # Click New Privilege
                    self.driver.find_element(By.LINK_TEXT, "New Privilege").click()
                    time.sleep(2)
                    
                    # Select application
                    select_element = Select(self.driver.find_element(By.NAME, "application_id"))
                    select_element.select_by_visible_text(app_name)
                    time.sleep(2)
                    
                    # Set Processing Status to "Complete"
                    print(f"  🔧 Setting Processing Status to 'Complete'...")
                    try:
                        # Click the Processing Status dropdown
                        status_dropdown = self.driver.find_element(By.XPATH, "//*[@id='status_id']")
                        status_dropdown.click()
                        print(f"  ✅ Clicked Processing Status dropdown")
                        time.sleep(0.5)
                        
                        # Click the second option (Complete)
                        complete_option = self.driver.find_element(By.XPATH, "//*[@id='status_id']/option[2]")
                        complete_option.click()
                        print(f"  ✅ Selected 'Complete' status")
                        time.sleep(1)
                    except Exception as e:
                        print(f"  ⚠️  Could not set Processing Status: {str(e)[:100]}")
                    
                    # Fill in "Performed By" field with autocomplete
                    print(f"  🔍 Looking for 'Performed By' field...")
                    performed_by_filled = False
                    performed_by_selectors = [
                        (By.ID, "searchField"),  # The actual ID from the Brown MyAccount page
                        (By.XPATH, "//*[@id='searchField']"),
                        (By.NAME, "performed_by"),
                        (By.NAME, "performedBy"),
                        (By.NAME, "performed_by_id"),
                        (By.ID, "performed_by"),
                        (By.ID, "performedBy"),
                        (By.XPATH, "//input[@placeholder='Performed By']"),
                        (By.XPATH, "//label[contains(text(), 'Performed By')]/following-sibling::input"),
                        (By.XPATH, "//label[contains(text(), 'Performed By')]/..//input"),
                    ]
                    
                    performed_by_field = None
                    for selector_type, selector_value in performed_by_selectors:
                        try:
                            performed_by_field = self.driver.find_element(selector_type, selector_value)
                            print(f"  ✅ Found 'Performed By' field using selector: {selector_value}")
                            break
                        except:
                            continue
                    
                    if not performed_by_field:
                        result['error'] = "⚠️ 'Performed By' field not found - stopping process"
                        print(f"  ❌ 'Performed By' field not found - tried {len(performed_by_selectors)} selectors")
                        results.append(result)
                        # Don't close driver - let user stay logged in
                        raise Exception("Performed By field not found. ChromeDriver session kept alive for debugging.")
                    
                    # Scroll the field into view before interacting with it
                    # This is important for autocomplete dropdowns to appear correctly
                    print(f"  📜 Scrolling 'Performed By' field into view...")
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", performed_by_field)
                        time.sleep(1)  # Wait for scroll to complete
                        print(f"  ✅ Scrolled field into view")
                    except Exception as e:
                        print(f"  ⚠️  Could not scroll field into view: {str(e)[:100]}")
                    
                    # Click the field first to focus it and activate autocomplete
                    try:
                        performed_by_field.click()
                        print(f"  👆 Clicked 'Performed By' field to focus it")
                        time.sleep(0.5)
                    except Exception as e:
                        print(f"  ⚠️  Could not click field: {str(e)[:100]}")
                    
                    # Type the name in the field
                    # Try normal send_keys first, fall back to JavaScript if it fails
                    try:
                        performed_by_field.send_keys(performed_by_name)
                        print(f"  ⌨️  Typed '{performed_by_name}' in Performed By field")
                    except Exception as e:
                        print(f"  ⚠️  Normal typing failed: {str(e)[:100]}")
                        print(f"  🔄 Trying JavaScript to set value and trigger events...")
                        try:
                            # Use JavaScript to set the value and trigger multiple events
                            self.driver.execute_script("""
                            var element = arguments[0];
                            var value = arguments[1];
                            element.value = value;
                            
                            // Trigger multiple events to activate autocomplete
                            element.dispatchEvent(new Event('focus', { bubbles: true }));
                            element.dispatchEvent(new Event('input', { bubbles: true }));
                            element.dispatchEvent(new Event('keydown', { bubbles: true }));
                            element.dispatchEvent(new Event('keyup', { bubbles: true }));
                            element.dispatchEvent(new Event('change', { bubbles: true }));
                            """, performed_by_field, performed_by_name)
                            print(f"  ✅ Set value via JavaScript and triggered events: '{performed_by_name}'")
                        except Exception as e2:
                            result['error'] = f"⚠️ Could not type in 'Performed By' field: {str(e2)[:100]}"
                            print(f"  ❌ JavaScript also failed: {str(e2)[:100]}")
                            results.append(result)
                            raise Exception(f"Could not type in Performed By field. ChromeDriver session kept alive for debugging.")
                    
                    # Wait 2 seconds for autocomplete suggestions to appear
                    print(f"  ⏱️  Waiting 2 seconds for autocomplete dropdown...")
                    time.sleep(2)

                    # Try clicking the autocomplete suggestion using XPath
                    print(f"  🔍 Looking for autocomplete suggestion...")
                    autocomplete_clicked = False
                    
                    # Try the specific XPath: //span/div[1]
                    try:
                        print(f"  🎯 Trying XPath: //span/div[1]...")
                        first_suggestion = WebDriverWait(self.driver, 3).until(
                            EC.element_to_be_clickable((By.XPATH, "//span/div[1]"))
                        )
                        
                        # Use ActionChains for real mouse click
                        actions = ActionChains(self.driver)
                        actions.move_to_element(first_suggestion).click().perform()
                        
                        print(f"  ✅ Clicked autocomplete suggestion")
                        autocomplete_clicked = True
                        performed_by_filled = True
                        
                    except Exception as e:
                        print(f"  ⚠️  XPath click failed: {str(e)[:100]}")
                    
                    if not autocomplete_clicked:
                        print(f"  🔍 Trying element-based clicking as fallback...")
                        autocomplete_selectors = [
                            # Try the exact visible dropdown structure first
                            (By.XPATH, "//div[contains(text(), 'Jameel')]"),  # Specific to the visible name
                            (By.XPATH, "//*[contains(@class, 'autocomplete')]//*[contains(text(), 'Jameel')]"),
                            
                            # Generic first item selectors
                            (By.XPATH, "//ul/li[1]"),  # First li in any ul
                            (By.XPATH, "//div[contains(@class, 'autocomplete')]//*[1]"),
                            (By.XPATH, "//ul[contains(@class, 'autocomplete')]//li[1]"),
                            (By.XPATH, "//div[contains(@class, 'suggestion')]//div[1]"),
                            (By.XPATH, "//div[contains(@class, 'suggestion')][1]"),
                            (By.XPATH, "//li[contains(@class, 'suggestion')][1]"),
                            
                            # Role-based selectors
                            (By.XPATH, "//div[@role='option'][1]"),
                            (By.XPATH, "//li[@role='option'][1]"),
                            (By.XPATH, "//*[@role='option'][1]"),
                            
                            # Common UI library patterns
                            (By.CSS_SELECTOR, ".autocomplete-suggestion:first-child"),
                            (By.CSS_SELECTOR, ".ui-menu-item:first-child"),
                            (By.XPATH, "//div[contains(@class, 'ui-menu-item')][1]"),
                            (By.XPATH, "//li[contains(@class, 'ui-menu-item')][1]"),
                            
                            # ID-based (in case there's a results list)
                            (By.XPATH, "//*[@id='searchFieldResults']//li[1]"),
                            (By.XPATH, "//*[@id='searchFieldResults']//*[1]"),
                            (By.XPATH, "//*[contains(@id, 'result')]//*[1]"),
                            
                            # Any visible div/li that contains the typed name
                            (By.XPATH, f"//*[contains(text(), '{performed_by_name.split()[0]}')]"),
                        ]
                        
                        for i, (selector_type, selector_value) in enumerate(autocomplete_selectors):
                            try:
                                print(f"     Trying selector {i+1}/{len(autocomplete_selectors)}: {selector_value[:80]}...")
                                first_suggestion = WebDriverWait(self.driver, 2).until(
                                    EC.presence_of_element_located((selector_type, selector_value))
                                )
                                
                                # Make sure element is visible
                                if not first_suggestion.is_displayed():
                                    print(f"     ⚠️  Element found but not visible")
                                    continue
                                
                                # Use ActionChains for a real mouse click instead of element.click()
                                print(f"     ✅ Found suggestion, performing real mouse click...")
                                actions = ActionChains(self.driver)
                                actions.move_to_element(first_suggestion).click().perform()
                                
                                print(f"  ✅ Clicked first autocomplete suggestion with real mouse click")
                                autocomplete_clicked = True
                                performed_by_filled = True
                                break
                            except Exception as e:
                                print(f"     ❌ Failed: {str(e)[:60]}")
                                continue
                    
                    if not autocomplete_clicked:
                        result['error'] = f"⚠️ No autocomplete suggestions appeared for '{performed_by_name}' - stopping process"
                        print(f"  ❌ No autocomplete suggestions found - tried {len(autocomplete_selectors)} selectors")
                        results.append(result)
                        # Don't close driver - let user stay logged in
                        raise Exception(f"No autocomplete suggestions appeared for '{performed_by_name}'. ChromeDriver session kept alive for debugging.")
                    
                    time.sleep(1)  # Wait for selection to register
                    
                    # If Employment Source is Banner: set attn_type and attn_date
                    if employment_source and "Banner" in employment_source:
                        print(f"  📋 Employment Source is Banner - setting attn_type and attn_date")
                        try:
                            attn_type_elem = self.driver.find_element(By.XPATH, "//*[@id='attn_type']")
                            attn_type_elem.click()
                            time.sleep(0.5)
                            attn_option = self.driver.find_element(By.XPATH, "//*[@id='attn_type']/option[3]")
                            attn_option.click()
                            print(f"  ✅ Selected attn_type option 3")
                            time.sleep(0.5)
                            
                            attn_date_str = self._get_next_attn_date()
                            attn_date_elem = self.driver.find_element(By.XPATH, "//*[@id='attn_date']")
                            attn_date_elem.clear()
                            attn_date_elem.send_keys(attn_date_str)
                            print(f"  ✅ Set attn_date to {attn_date_str}")
                            time.sleep(0.5)
                        except Exception as e:
                            print(f"  ⚠️  Could not set attn_type/attn_date (Banner): {str(e)[:100]}")
                    else:
                        print(f"  📋 Employment Source is not Banner - skipping attn_type/attn_date")
                    
                    # Add comment
                    textarea = self.driver.find_element(By.NAME, "comments")
                    existing_text = textarea.get_attribute("value") or ""
                    new_text = (existing_text + " " + comment).strip() if comment else existing_text
                    textarea.clear()
                    textarea.send_keys(new_text)
                    
                    # Save
                    submit_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//button[@type="submit" and text()="Save"]'))
                    )
                    submit_button.click()
                    time.sleep(2)
                    
                    # Return to People page
                    people_button = self.driver.find_element(
                        By.XPATH,
                        "//a[@class='selected' and contains(text(), 'People')]"
                    )
                    people_button.click()
                    
                    result['success'] = True
                except OperationAbortedException:
                    raise
                except Exception as e:
                    result['error'] = str(e)
                
                results.append(result)
        except OperationAbortedException:
            print("⛔ Add privileges aborted by user")
        
        return results
    
    def revoke_privileges(self, ids, app_name, dp_number):
        """Revoke privileges from a list of user IDs"""
        if not self.driver:
            raise Exception("Not logged in. Please login first.")
        
        results = []
        
        try:
            for user_id in ids:
                self._check_abort()
                result = {'id': user_id, 'success': False, 'error': None}
                try:
                    # Search for user
                    search_button = self.driver.find_element(By.NAME, "search")
                    text_box = self.driver.find_element(By.NAME, "brown_login")
                    text_box.clear()
                    text_box.send_keys(user_id)
                    time.sleep(3)
                    search_button.click()
                    time.sleep(3)
                    
                    # Click View Overview
                    try:
                        vo = self.driver.find_element(
                            By.XPATH,
                            "//a[@class='btn btn-default' and contains(text(), 'View Overview')]"
                        )
                        vo.click()
                    except:
                        result['error'] = "Could not find View Overview"
                        results.append(result)
                        self.driver.find_element(
                            By.XPATH,
                            "//a[@class='selected' and contains(text(), 'People')]"
                        ).click()
                        time.sleep(3)
                        continue
                    
                    time.sleep(2)
                    
                    # Navigate to AdminID - Current
                    try:
                        button = self.driver.find_element(By.LINK_TEXT, "AdminID - Current")
                        button.click()
                    except:
                        result['error'] = "Could not find AdminID - Current"
                        results.append(result)
                        self.driver.find_element(
                            By.XPATH,
                            "//a[@class='selected' and contains(text(), 'People')]"
                        ).click()
                        time.sleep(3)
                        continue
                    
                    # Find privilege in table - iterate over TR elements (direct rows), not inside a single tr
                    try:
                        tbody = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.TAG_NAME, 'tbody'))
                        )
                        time.sleep(1)
                        rows = tbody.find_elements(By.XPATH, './tr')
                    except Exception:
                        result['error'] = "No privileges found"
                        results.append(result)
                        self.driver.find_element(
                            By.XPATH,
                            "//a[@class='selected' and contains(text(), 'People')]"
                        ).click()
                        time.sleep(3)
                        continue
                    
                    privilege_found = False
                    for i in range(len(rows)):
                        try:
                            rows = tbody.find_elements(By.XPATH, './tr')
                            if i >= len(rows):
                                break
                            row = rows[i]
                            spans = row.find_elements(By.XPATH, './/span')
                            print(f"  [DEBUG] Row {i + 1}/{len(rows)}: checking {len(spans)} span(s) for '{app_name}'")
                            for j, s in enumerate(spans):
                                raw_text = (s.text or "").strip()
                                data_content = s.get_attribute("data-content") or ""
                                print(f"    span[{j}] raw_text={repr(s.text)} -> trimmed={repr(raw_text)} | data-content={repr(data_content)} | match_text={app_name in raw_text} | match_data={app_name in data_content}")
                            # Match by element content: span text (e.g. "BIOR") or data-content (e.g. "BioRender")
                            try:
                                span = row.find_element(By.XPATH, f'.//span[contains(normalize-space(.), "{app_name}") or contains(@data-content, "{app_name}")]')
                            except NoSuchElementException:
                                continue
                            # Click Edit
                            edit_link = row.find_element(By.XPATH, './/a[contains(text(), "Edit")]')
                            edit_link.click()
                            time.sleep(3)
                            
                            # Check current expiration reason; if already set, skip and note it
                            select_element = Select(self.driver.find_element(By.NAME, "exp_reason"))
                            current_reason_elem = select_element.first_selected_option
                            current_reason = (current_reason_elem.text or "").strip()
                            if current_reason and current_reason.lower() != "select a reason":
                                result['error'] = f"Privilege already {current_reason}"
                                result['success'] = True
                                self.driver.find_element(By.LINK_TEXT, "People").click()
                                time.sleep(3)
                                privilege_found = True
                                break
                            
                            # Set expiration reason
                            select_element.select_by_visible_text("Revoked")
                            time.sleep(1)
                            
                            # Add comment (append on new line, don't overwrite existing content)
                            textarea = None
                            try:
                                textarea = self.driver.find_element(By.ID, "comments")
                            except Exception:
                                textarea = self.driver.find_element(By.NAME, "comments")
                            existing_text = textarea.get_attribute("value") or ""
                            new_text = (existing_text.rstrip() + "\n" + dp_number).strip() if dp_number else existing_text
                            textarea.clear()
                            textarea.send_keys(new_text)
                            time.sleep(1)
                            
                            # Click save button (must click before returning to search)
                            save_xpath = "/html/body/div[1]/div/div[2]/div[3]/form/div[2]/div/div/div/button"
                            submit_button = None
                            try:
                                submit_button = WebDriverWait(self.driver, 10).until(
                                    EC.presence_of_element_located((By.XPATH, save_xpath))
                                )
                            except Exception:
                                try:
                                    submit_button = self.driver.find_element(By.XPATH, "//form//button[@type='submit']")
                                except Exception:
                                    try:
                                        submit_button = self.driver.find_element(By.XPATH, "//form//button[contains(., 'Save')]")
                                    except Exception:
                                        pass
                            if not submit_button:
                                result['error'] = "Save button not found"
                                self.driver.find_element(By.LINK_TEXT, "People").click()
                                time.sleep(3)
                                break
                            self.driver.execute_script("arguments[0].click();", submit_button)
                            time.sleep(3)
                            
                            # Return to People page
                            self.driver.find_element(By.LINK_TEXT, "People").click()
                            time.sleep(3)
                            
                            result['success'] = True
                            privilege_found = True
                            break
                        except Exception as e:
                            result['error'] = str(e)
                            try:
                                self.driver.find_element(By.LINK_TEXT, "People").click()
                                time.sleep(3)
                            except Exception:
                                pass
                            break
                    
                    if not privilege_found:
                        if not result.get('error'):
                            result['error'] = f"Privilege for '{app_name}' not found"
                        try:
                            self.driver.find_element(By.LINK_TEXT, "People").click()
                            time.sleep(3)
                        except Exception:
                            pass
                
                except OperationAbortedException:
                    raise
                except Exception as e:
                    result['error'] = str(e)
                
                results.append(result)
        except OperationAbortedException:
            print("⛔ Revoke privileges aborted by user")
        
        return results
    
    def get_employment_status(self, ids, id_type='SID', to_fields=None, on_result_callback=None):
        """
        Get status fields: Source System, Employment Status, Student Status Code.
        to_fields: list of 'SOURCE_SYSTEM', 'EMPLOYMENT_STATUS', 'STUDENT_STATUS_CODE'
        """
        if not self.driver:
            raise Exception("Not logged in. Please login first.")
        if not to_fields:
            to_fields = ['SOURCE_SYSTEM', 'EMPLOYMENT_STATUS']
        results = []
        try:
            for index, user_id in enumerate(ids):
                self._check_abort()
                result = {
                    'id': user_id, 'source': None, 'employment_status': None, 'student_status_code': None,
                    'success': False, 'error': None
                }
                try:
                    search_button = self.driver.find_element(By.NAME, "search")
                    if id_type == 'BID':
                        text_box = self.driver.find_element(By.ID, "brown_id")
                    elif id_type == 'SID':
                        text_box = self.driver.find_element(By.NAME, "brown_login")
                    else:
                        text_box = self.driver.find_element(By.NAME, "brown_login")
                    text_box.clear()
                    text_box.send_keys(user_id)
                    time.sleep(2)
                    search_button.click()
                    time.sleep(3)
                    try:
                        vo = self.driver.find_element(
                            By.XPATH,
                            "//a[@class='btn btn-default' and contains(text(), 'View Overview')]"
                        )
                        vo.click()
                    except Exception:
                        result['error'] = "User not found"
                        results.append(result)
                        if on_result_callback:
                            on_result_callback(index, result)
                        continue
                    time.sleep(2)
                    self._check_abort()
                    try:
                        # Same Overview page for BID and SID; use same XPaths as Short ID
                        employment_status_div = self.driver.find_element(
                            By.XPATH,
                            "//b[@class='col-xs-6' and contains(text(), 'Employment Status:')]/following-sibling::div[@class='col-xs-6']"
                        )
                        source_div = self.driver.find_element(
                            By.XPATH,
                            "//b[@class='col-xs-6' and contains(text(), 'Source')]/following-sibling::div[@class='col-xs-6']"
                        )
                        if 'SOURCE_SYSTEM' in to_fields:
                            result['source'] = source_div.text.strip()
                        if 'EMPLOYMENT_STATUS' in to_fields:
                            result['employment_status'] = employment_status_div.text.strip()
                        result['success'] = True
                    except Exception as e:
                        result['error'] = f"Could not extract data: {str(e)}"
                    if result['success'] and 'STUDENT_STATUS_CODE' in to_fields:
                        student_status_xpath = "/html/body/div[1]/div/div[2]/div[3]/div[2]/div/div[4]/div[1]/div/div/div"
                        result['student_status_code'] = ''
                        for link_index in (4, 5):
                            try:
                                link = self.driver.find_element(
                                    By.XPATH,
                                    f"/html/body/div[1]/div/div[1]/div/div/a[{link_index}]"
                                )
                                link.click()
                                time.sleep(2)
                                self._check_abort()
                                student_div = self.driver.find_element(By.XPATH, student_status_xpath)
                                value = student_div.text.strip()
                                if value:
                                    result['student_status_code'] = value
                                    break
                            except Exception as e:
                                if link_index == 4:
                                    print(f"  Student Status Code not at a[4], trying a[5]: {e}")
                                else:
                                    print(f"  ⚠️  Student Status Code not found: {e}")
                                continue
                    people_button = self.driver.find_element(
                        By.XPATH,
                        "//a[@class='selected' and contains(text(), 'People')]"
                    )
                    people_button.click()
                except OperationAbortedException:
                    raise
                except Exception as e:
                    result['error'] = str(e)
                
                results.append(result)
                
                # Call callback immediately after result
                if on_result_callback:
                    on_result_callback(index, result)
        except OperationAbortedException:
            print("⛔ Employment status check aborted by user")
        
        return results
    
    def convert_ids(self, ids, from_type, to_types, write_columns, on_result_callback=None):
        """
        Convert between ID types. Can output multiple types at once from one Overview view.
        After View Overview, all target IDs are read from the Overview page once.
        
        Args:
            ids: List of IDs to convert
            from_type: Source ID type ('NETID', 'BID', 'BROWN_EMAIL', or 'SID')
            to_types: List of target types e.g. ['SID', 'NETID']
            write_columns: Dict mapping to_type -> column letter e.g. {'SID': 'B', 'NETID': 'C'}
            on_result_callback: Optional function to call after each result with (index, result)
        """
        if not self.driver:
            raise Exception("Not logged in. Please login first.")
        
        valid_to_types = ('SID', 'NETID', 'BID', 'BROWN_EMAIL')
        if not to_types:
            raise Exception("At least one target type (to_types) is required")
        for t in to_types:
            if t not in valid_to_types:
                raise Exception(f"to_type must be one of: {', '.join(valid_to_types)}")
        
        results = []
        
        try:
            for index, user_id in enumerate(ids):
                self._check_abort()
                result = {'id': user_id, 'converted_ids': {}, 'success': False, 'error': None}
                try:
                    search_button = self.driver.find_element(By.NAME, "search")
                    
                    if from_type == 'NETID':
                        netid = user_id.replace("@brown.edu", "")
                        text_box = self.driver.find_element(By.NAME, "brown_netid")
                        text_box.clear()
                        text_box.send_keys(netid)
                    elif from_type == 'BID':
                        text_box = self.driver.find_element(By.NAME, "brown_id")
                        text_box.clear()
                        text_box.send_keys(user_id)
                    elif from_type == 'BROWN_EMAIL':
                        netid = user_id.split('@')[0].strip()
                        text_box = self.driver.find_element(By.NAME, "brown_netid")
                        text_box.clear()
                        text_box.send_keys(netid)
                    elif from_type == 'SID':
                        text_box = self.driver.find_element(By.NAME, "brown_login")
                        text_box.clear()
                        text_box.send_keys(user_id)
                    else:
                        result['error'] = f"Unsupported from_type: {from_type}"
                        results.append(result)
                        if on_result_callback:
                            on_result_callback(index, result)
                        continue
                    
                    time.sleep(2)
                    search_button.click()
                    time.sleep(2)
                    
                    try:
                        vo = self.driver.find_element(
                            By.XPATH,
                            "//a[@class='btn btn-default' and contains(text(), 'View Overview')]"
                        )
                        vo.click()
                    except:
                        result['error'] = "User not found"
                        results.append(result)
                        if on_result_callback:
                            on_result_callback(index, result)
                        continue
                    
                    time.sleep(2)
                    
                    # Read all ID types from Overview page once
                    netid_elem = self.driver.find_element(
                        By.XPATH,
                        "/html/body/div[1]/div/div[2]/div[3]/div[1]/div/div[7]/div[1]/div/div/div"
                    )
                    sid_elem = self.driver.find_element(
                        By.XPATH,
                        "/html/body/div[1]/div/div[2]/div[3]/div[1]/div/div[7]/div[2]/div/div"
                    )
                    bid_elem = self.driver.find_element(
                        By.XPATH,
                        "/html/body/div[1]/div/div[2]/div[3]/div[1]/div/div[1]/div[1]/div/div"
                    )
                    netid_val = netid_elem.text.strip()
                    sid_val = sid_elem.text.strip()
                    bid_val = bid_elem.text.strip()
                    brown_email_val = f"{netid_val}@brown.edu"
                    values = {'SID': sid_val, 'NETID': netid_val, 'BID': bid_val, 'BROWN_EMAIL': brown_email_val}
                    for t in to_types:
                        result['converted_ids'][t] = values.get(t, '')
                    result['success'] = True
                    
                    people_button = self.driver.find_element(
                        By.XPATH,
                        "//a[@class='selected' and contains(text(), 'People')]"
                    )
                    people_button.click()
                    
                except OperationAbortedException:
                    raise
                except Exception as e:
                    result['error'] = str(e)
                
                results.append(result)
                if on_result_callback:
                    on_result_callback(index, result)
        except OperationAbortedException:
            print("⛔ ID conversion aborted by user")
        
        return results
    
    def logout(self):
        """Close browser and cleanup"""
        self._cleanup_driver()
    
    def __del__(self):
        """Cleanup when object is destroyed"""
        self._cleanup_driver()

