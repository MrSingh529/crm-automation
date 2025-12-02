import asyncio
import json
import logging
import sys
import os
from datetime import datetime
from pathlib import Path
import pandas as pd
from typing import List, Optional, Dict, Any
from playwright.async_api import async_playwright, Page, Browser, BrowserContext

# Configure logging
def setup_logging(log_level: str = "INFO"):
    """Setup logging configuration"""
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"automation_{timestamp}.log"
    
    logging.basicConfig(
        level=getattr(logging, log_level),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    return logging.getLogger(__name__)

class CRMAutomation:
    def __init__(self, config_path: str = "config.json", excel_path: Optional[str] = None):
        """Initialize CRM Automation"""
        # FIRST: Setup logger
        self.logger = setup_logging("INFO")  # Default to INFO initially
        
        # THEN: Load config
        self.config = self.load_config(config_path)
        
        # Update logger with config level
        log_level = self.config.get("log_level", "INFO")
        self.logger.setLevel(getattr(logging, log_level))
        
        self.excel_path = excel_path or self.config.get("excel_path")
        
        # Setup paths
        self.download_dir = Path(self.config.get("download_path", "invoices"))
        self.download_dir.mkdir(exist_ok=True)
        
        # Statistics
        self.stats = {
            "total": 0,
            "success": 0,
            "failed": 0,
            "skipped": 0,
            "start_time": None,
            "end_time": None
        }
        
        self.cancelled = False

        # YOUR CRM XPATHS - UPDATED WITH YOUR SPECIFIC PATHS
        self.xpaths = {
            # Login page
            "username": "xpath=/html/body/div/form/div[2]/input",
            "password": "xpath=/html/body/div/form/div[3]/input",
            "login_button": "xpath=/html/body/div/form/div[3]/button",
            "logout_button": "xpath=/html/body/div/div/div[1]/div[2]/ul/li[10]/a",
            
            # Job Search navigation
            "job_search_menu": "xpath=/html/body/div[1]/div/div[1]/div[2]/ul/li[4]",  # First click to expand
            "job_search_link": "xpath=/html/body/div[1]/div/div[1]/div[2]/ul/ul[3]/a/li",  # Second click
            
            # Search page
            "search_input": "xpath=/html/body/div/div/div[2]/form[2]/div/div/div[2]/label/input",
            
            # Results table
            "eye_button": "xpath=/html/body/div/div/div[2]/form[2]/div/div/table/tbody/tr/td[16]/div/a/i",
            "table_rows": "xpath=/html/body/div/div/div[2]/form[2]/div/div/table/tbody/tr",
            
            # Details page
            "invoice_image": "xpath=/html/body/div/div/div[2]/form/div[1]/div[2]/div[2]/table/tbody/tr[8]/td[2]/a",
        }
        
        self.logger.info("CRM Automation initialized with custom XPaths")
    
    async def check_cancelled(self):
        """Check if automation was cancelled"""
        if self.cancelled:
            self.logger.info("Automation cancelled by user")
            raise KeyboardInterrupt("Automation cancelled")

    async def smart_wait(self, page: Page, timeout_seconds: int = 30):
        """Smart waiting for page to be ready"""
        self.logger.debug(f"Smart wait for {timeout_seconds} seconds...")
        
        start_time = datetime.now()
        timeout_ms = timeout_seconds * 1000
        
        try:
            # Wait for network to be idle
            await page.wait_for_load_state("networkidle", timeout=timeout_ms)
            
            # Wait a bit more for JavaScript to execute
            await page.wait_for_timeout(2000)
            
            # Check if page has content
            body_text = await page.text_content("body")
            if body_text and len(body_text.strip()) > 100:  # Reasonable amount of content
                self.logger.debug("✓ Page has content")
            else:
                self.logger.warning("Page seems empty, waiting more...")
                await page.wait_for_timeout(3000)
                
        except Exception as e:
            self.logger.warning(f"Wait interrupted: {e}")
        
        elapsed = (datetime.now() - start_time).total_seconds()
        self.logger.debug(f"Waited {elapsed:.1f} seconds")

    def load_config(self, config_path: str) -> Dict[str, Any]:
        """Load configuration from JSON file"""
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
            self.logger.info(f"Configuration loaded from {config_path}")
            return config
        except FileNotFoundError:
            print(f"ERROR: Configuration file not found: {config_path}")
            print("Please run app.py first to create configuration.")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"ERROR: Invalid JSON in config file: {e}")
            sys.exit(1)
    
    def load_service_orders(self) -> List[str]:
        """Load Service Order numbers from Excel file"""
        try:
            df = pd.read_excel(self.excel_path)
            
            # Find the Service Order column (case-insensitive)
            so_columns = [col for col in df.columns if any(
                keyword in str(col).lower() 
                for keyword in ['service order', 'so', 'order no', 'job no', 'job no.']
            )]
            
            if not so_columns:
                raise ValueError("No Service Order column found in Excel file")
            
            so_column = so_columns[0]
            self.logger.info(f"Using column '{so_column}' for Service Order numbers")
            
            # Extract and clean SO numbers
            so_numbers = df[so_column].dropna().astype(str).str.strip().tolist()
            
            if not so_numbers:
                raise ValueError("No Service Order numbers found in the Excel file")
            
            self.logger.info(f"Loaded {len(so_numbers)} Service Order numbers")
            return so_numbers
            
        except Exception as e:
            self.logger.error(f"Failed to load Service Orders: {e}")
            raise
    
    async def setup_browser(self) -> tuple:
        """Setup Playwright browser and context"""
        playwright = await async_playwright().start()
        
        launch_args = {
            "headless": self.config.get("headless", False),
            "args": [
                "--start-maximized", 
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ],
            "slow_mo": 100  # SLOW DOWN actions by 100ms for visibility
        }
        
        browser = await playwright.chromium.launch(**launch_args)
        
        # Configure context with download support
        context_args = {
            "viewport": {"width": 1920, "height": 1080},  # Fixed size instead of None
            "accept_downloads": True,
            "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "ignore_https_errors": True,
            "bypass_csp": True  # Bypass Content Security Policy if needed
        }
        
        # Load storage state if exists (for persistent login)
        storage_state = self.config.get("storage_state")
        if storage_state and Path(storage_state).exists():
            context_args["storage_state"] = storage_state
        
        context = await browser.new_context(**context_args)
        
        # Set default timeout
        context.set_default_timeout(45000)  # Increased to 45 seconds
        context.set_default_navigation_timeout(60000)  # 60 seconds for navigation
        
        page = await context.new_page()
        
        # Add page event listeners for debugging
        page.on("load", lambda: self.logger.debug("Page loaded"))
        page.on("domcontentloaded", lambda: self.logger.debug("DOM content loaded"))
        
        return playwright, browser, context, page
    
    async def login(self, page: Page) -> bool:
        """Login to CRM system using YOUR XPaths"""
        try:
            self.logger.info("Logging into CRM...")
            
            crm_url = self.config["crm_url"]
            username = self.config["username"]
            password = self.config["password"]
            
            # Go to CRM with longer timeout
            await page.goto(crm_url, wait_until="domcontentloaded", timeout=60000)
            
            # Wait for page to be fully loaded
            await page.wait_for_load_state("networkidle", timeout=60000)
            
            # ADDED: Give extra time for page to render
            await page.wait_for_timeout(5000)  # Wait 5 seconds for page to fully render
            
            # Wait for login page to load - with multiple attempts
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    self.logger.info(f"Looking for username field (Attempt {attempt + 1}/{max_attempts})...")
                    
                    # Try to find username field
                    username_field = await page.wait_for_selector(
                        self.xpaths["username"], 
                        timeout=10000,
                        state="visible"
                    )
                    
                    if username_field:
                        self.logger.info("✓ Found username field")
                        
                        # Fill credentials using YOUR XPaths
                        await username_field.fill(username)
                        
                        # Find and fill password
                        password_field = await page.wait_for_selector(
                            self.xpaths["password"], 
                            timeout=5000,
                            state="visible"
                        )
                        
                        if password_field:
                            await password_field.fill(password)
                            
                            # Click login button
                            login_button = await page.wait_for_selector(
                                self.xpaths["login_button"],
                                timeout=5000,
                                state="visible"
                            )
                            
                            if login_button:
                                self.logger.info("✓ Clicking login button...")
                                await login_button.click()
                                
                                # Wait for navigation - give more time
                                await page.wait_for_load_state("networkidle", timeout=30000)
                                await page.wait_for_timeout(5000)  # Extra wait for CRM
                                
                                # Check if login was successful
                                if await self.is_logged_in(page):
                                    self.logger.info("✓ Login successful")
                                    
                                    # Save storage state for future sessions
                                    await page.context.storage_state(path="auth_state.json")
                                    return True
                                else:
                                    self.logger.warning("Login may have failed - checking...")
                                    # Take screenshot for debugging
                                    await page.screenshot(path="login_attempt.png")
                                    continue  # Try again
                        
                except Exception as e:
                    self.logger.warning(f"Attempt {attempt + 1} failed: {e}")
                    await page.wait_for_timeout(3000)  # Wait before retry
                    
                    # Try reloading the page
                    if attempt < max_attempts - 1:
                        await page.reload(wait_until="networkidle", timeout=30000)
                        await page.wait_for_timeout(3000)
            
            self.logger.error("All login attempts failed")
            await page.screenshot(path="login_failed.png")
            return False
            
        except Exception as e:
            self.logger.error(f"Login failed with error: {e}")
            await page.screenshot(path="login_error.png")
            return False
    
    async def is_logged_in(self, page: Page) -> bool:
        """Check if user is logged in using multiple methods"""
        try:
            # Method 1: Check for logout button using YOUR XPath
            try:
                logout_button = await page.wait_for_selector(
                    self.xpaths["logout_button"], 
                    timeout=5000,
                    state="visible"
                )
                if logout_button:
                    self.logger.debug("✓ Found logout button - user is logged in")
                    return True
            except:
                pass
            
            # Method 2: Check for URL change (not on login page)
            current_url = page.url
            if "login" not in current_url.lower() and "signin" not in current_url.lower():
                self.logger.debug("✓ Not on login page - likely logged in")
                
                # Method 3: Check for dashboard elements
                dashboard_selectors = [
                    'div.dashboard', 'nav', 'header', '.main-content',
                    '.sidebar', '.menu', '#navbar', '.top-bar'
                ]
                
                for selector in dashboard_selectors:
                    elements = await page.query_selector_all(selector)
                    if len(elements) > 0:
                        self.logger.debug(f"✓ Found dashboard element: {selector}")
                        return True
            
            # Method 4: Check for welcome message or username display
            welcome_selectors = ['text=/welcome/i', 'text=/hello/i', '.user-name', '.profile-name']
            for selector in welcome_selectors:
                try:
                    element = await page.query_selector(selector)
                    if element:
                        self.logger.debug(f"✓ Found welcome element: {selector}")
                        return True
                except:
                    pass
            
            self.logger.debug("✗ Could not confirm login status")
            return False
            
        except Exception as e:
            self.logger.debug(f"Login check error: {e}")
            return False
    
    async def navigate_to_job_search(self, page: Page) -> bool:
        """Navigate to Job Search page using YOUR XPaths"""
        try:
            self.logger.info("Navigating to Job Search...")
            
            # First click: Expand Job Search menu
            await page.click(self.xpaths["job_search_menu"])
            await page.wait_for_timeout(1000)  # Wait for menu to expand
            
            # Second click: Click Job Search link
            await page.click(self.xpaths["job_search_link"])
            
            # Wait for page to load
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(3000)  # Extra wait for CRM
            
            # Verify we're on the right page by checking search input
            try:
                await page.wait_for_selector(self.xpaths["search_input"], timeout=10000)
                self.logger.info("Successfully navigated to Job Search page")
                return True
            except:
                self.logger.warning("Search input not found, but continuing...")
                return True  # Continue anyway
                
        except Exception as e:
            self.logger.error(f"Failed to navigate to Job Search: {e}")
            await page.screenshot(path="navigation_error.png")
            return False
    
    async def search_service_order(self, page: Page, so_number: str) -> bool:
        """Search for a specific Service Order"""
        try:
            self.logger.debug(f"Searching for SO: {so_number}")
            
            # Wait for search input using YOUR XPath
            try:
                search_box = await page.wait_for_selector(self.xpaths["search_input"], timeout=10000)
            except:
                self.logger.error("Search box not found")
                return False
            
            # Clear and enter search term
            await search_box.fill('')
            await search_box.type(so_number, delay=100)
            await page.keyboard.press("Enter")
            
            # Wait for results - longer wait for CRM
            await page.wait_for_timeout(4000)
            await page.wait_for_load_state("networkidle")
            
            # Check if results are found
            results_found = await self.verify_search_results(page, so_number)
            
            if results_found:
                self.logger.debug(f"Search successful for {so_number}")
                return True
            else:
                self.logger.warning(f"No results found for {so_number}")
                return False
                
        except Exception as e:
            self.logger.error(f"Search failed for {so_number}: {e}")
            return False
    
    async def verify_search_results(self, page: Page, so_number: str) -> bool:
        """Verify that search returned results"""
        try:
            # Wait for table rows using YOUR XPath pattern
            try:
                await page.wait_for_selector(self.xpaths["table_rows"], timeout=15000)
            except:
                self.logger.debug(f"No table rows found for {so_number}")
                return False
            
            # Get all rows
            rows = await page.query_selector_all(self.xpaths["table_rows"])
            
            if not rows:
                return False
            
            # Check if SO number appears in any row
            for row in rows:
                row_text = await row.text_content()
                if so_number in row_text:
                    self.logger.debug(f"Found SO {so_number} in results")
                    return True
            
            # Check for "no results" message
            no_results = await page.query_selector('text=/no.*results|not found|no data/i')
            if no_results:
                self.logger.debug(f"No results message found for {so_number}")
                return False
            
            self.logger.debug(f"SO {so_number} not found in {len(rows)} rows")
            return False
            
        except Exception as e:
            self.logger.debug(f"Verify search error: {e}")
            return False
    
    async def find_eye_button_for_so(self, page: Page, so_number: str):
        """Find the eye button for a specific Service Order"""
        try:
            # Get all rows
            rows = await page.query_selector_all(self.xpaths["table_rows"])
            
            for row in rows:
                row_text = await row.text_content()
                if so_number in row_text:
                    # Find eye button in this row - using XPath (add xpath= prefix)
                    eye_button = await row.query_selector("xpath=./td[16]/div/a/i")
                    if not eye_button:
                        # Try alternative: look for any icon/button in the last column (CSS selector)
                        eye_button = await row.query_selector("td:last-child i, td:last-child a, td:last-child button")
                    
                    if eye_button:
                        return eye_button
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error finding eye button: {e}")
            return None
    
    async def open_service_order_details(self, page: Page, so_number: str) -> bool:
        """Open Service Order details by clicking the eye button"""
        try:
            # Find eye button for this specific SO
            eye_button = await self.find_eye_button_for_so(page, so_number)
            
            if not eye_button:
                self.logger.warning(f"View button not found for {so_number}")
                return False
            
            # Scroll to the button if needed
            await eye_button.scroll_into_view_if_needed()
            await page.wait_for_timeout(1000)
            
            # Click the eye button
            await eye_button.click()
            
            # Wait as specified in requirements
            wait_time = self.config.get("wait_time", 5)
            self.logger.debug(f"Waiting {wait_time} seconds for details page...")
            await page.wait_for_timeout(wait_time * 1000)
            
            # Wait for details page to load
            await page.wait_for_load_state("networkidle")
            
            # Check if we're on details page by looking for invoice image
            try:
                await page.wait_for_selector(self.xpaths["invoice_image"], timeout=10000, state="visible")
                return True
            except:
                self.logger.debug(f"Invoice image not immediately found, but continuing...")
                return True  # Continue anyway
                
        except Exception as e:
            self.logger.error(f"Failed to open details for {so_number}: {e}")
            await page.screenshot(path=f"details_error_{so_number}.png")
            return False
    
    async def save_invoice_image_new_tab(self, page: Page, so_number: str) -> bool:
        """Handle invoice image that opens in new tab"""
        try:
            # Get the context to handle new tabs
            context = page.context
            
            # Wait for the invoice image link
            invoice_link = await page.wait_for_selector(
                self.xpaths["invoice_image"], 
                timeout=10000,
                state="visible"
            )
            
            if not invoice_link:
                self.logger.warning(f"No invoice image link found for {so_number}")
                return False
            
            # Click the link (will open in new tab)
            async with context.expect_page() as new_page_info:
                await invoice_link.click()
            
            new_page = await new_page_info.value
            
            # Wait for new page to load
            await new_page.wait_for_load_state("networkidle")
            await new_page.wait_for_timeout(2000)
            
            # Save the entire page as image
            save_path = self.download_dir / f"{so_number}.png"
            await new_page.screenshot(path=save_path, full_page=True)
            
            self.logger.info(f"✅ Invoice saved for {so_number}")
            
            # Close the new tab
            await new_page.close()
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to save invoice in new tab for {so_number}: {e}")
            return False
    
    async def find_and_save_invoice_image(self, page: Page, so_number: str) -> bool:
        """Find and save the invoice image"""
        try:
            # Scroll to make sure element is visible
            invoice_link = await page.query_selector(self.xpaths["invoice_image"])
            if invoice_link:
                await invoice_link.scroll_into_view_if_needed()
                await page.wait_for_timeout(1000)
            
            # Try new tab method first (based on your description)
            if await self.save_invoice_image_new_tab(page, so_number):
                return True
            
            # Fallback: try direct download if it's an image element
            image_element = await page.query_selector(f'{self.xpaths["invoice_image"]} img')
            if image_element:
                image_src = await image_element.get_attribute('src')
                if image_src:
                    if image_src.startswith('data:'):
                        # Base64 encoded image
                        await self.save_base64_image(image_src, so_number)
                        return True
                    elif image_src.startswith('http'):
                        # Direct image URL
                        await page.goto(image_src)
                        await page.screenshot(path=self.download_dir / f"{so_number}.png", full_page=True)
                        await page.go_back()
                        return True
            
            # Last resort: screenshot the area
            invoice_element = await page.query_selector(self.xpaths["invoice_image"])
            if invoice_element:
                await invoice_element.screenshot(path=self.download_dir / f"{so_number}.png")
                self.logger.info(f"✅ Invoice saved (screenshot) for {so_number}")
                return True
            
            self.logger.warning(f"No invoice image found for {so_number}")
            return False
            
        except Exception as e:
            self.logger.error(f"Failed to save invoice for {so_number}: {e}")
            await page.screenshot(path=f"invoice_error_{so_number}.png")
            return False
    
    async def save_base64_image(self, base64_string: str, so_number: str):
        """Save base64 encoded image"""
        import base64
        
        try:
            # Extract base64 data
            if 'base64,' in base64_string:
                base64_data = base64_string.split('base64,')[1]
            else:
                base64_data = base64_string
            
            # Decode and save
            image_data = base64.b64decode(base64_data)
            
            # Determine file extension
            if base64_string.startswith('data:image/png'):
                ext = 'png'
            elif base64_string.startswith('data:image/jpeg') or base64_string.startswith('data:image/jpg'):
                ext = 'jpg'
            else:
                ext = 'png'
            
            with open(self.download_dir / f"{so_number}.{ext}", 'wb') as f:
                f.write(image_data)
                
        except Exception as e:
            self.logger.error(f"Failed to save base64 image: {e}")
            raise
    
    async def return_to_search(self, page: Page) -> bool:
        """Return to search page for next iteration"""
        try:
            # Based on your description: click on job search again (second one)
            await page.click(self.xpaths["job_search_link"])
            await page.wait_for_load_state("networkidle")
            await page.wait_for_timeout(2000)
            
            # Verify we're back on search page
            if await page.query_selector(self.xpaths["search_input"]):
                return True
            return False
            
        except Exception as e:
            self.logger.warning(f"Failed to return to search: {e}. Trying alternative...")
            
            # Alternative: go back in history
            try:
                await page.go_back()
                await page.wait_for_load_state("networkidle")
                return True
            except:
                # Last resort: reload the job search URL
                await page.goto(f"{self.config['crm_url']}/job-search", wait_until="networkidle")
                return True
    
    async def process_service_order(self, page: Page, so_number: str, retry_count: int = 0) -> bool:
        """Process a single Service Order"""
        # Check if cancelled
        if self.cancelled:
            return False
        
        max_retries = self.config.get("max_retries", 3)
        
        try:
            self.logger.info(f"Processing SO: {so_number} (Attempt {retry_count + 1}/{max_retries})")
            
            # Step 1: Search for SO
            search_success = await self.search_service_order(page, so_number)
            if not search_success:
                self.logger.warning(f"Search failed for {so_number}")
                return False
            
            # Step 2: Open details
            details_opened = await self.open_service_order_details(page, so_number)
            if not details_opened:
                self.logger.warning(f"Failed to open details for {so_number}")
                return False
            
            # Step 3: Save invoice image
            invoice_saved = await self.find_and_save_invoice_image(page, so_number)
            if not invoice_saved:
                self.logger.warning(f"Failed to save invoice for {so_number}")
                return False
            
            self.logger.info(f"✅ Successfully processed {so_number}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error processing {so_number}: {e}")
            
            if retry_count < max_retries - 1:
                self.logger.info(f"Retrying {so_number}...")
                await page.wait_for_timeout(3000)
                return await self.process_service_order(page, so_number, retry_count + 1)
            else:
                return False
    
    async def run_automation(self):
        """Main automation loop"""
        self.stats["start_time"] = datetime.now()
        
        try:
            # Load Service Orders
            so_numbers = self.load_service_orders()
            self.stats["total"] = len(so_numbers)
            
            # Setup browser
            playwright, browser, context, page = await self.setup_browser()
            
            try:
                # Login
                login_success = await self.login(page)
                if not login_success:
                    self.logger.error("Login failed. Exiting.")
                    return
                
                # Navigate to Job Search
                navigation_success = await self.navigate_to_job_search(page)
                if not navigation_success:
                    self.logger.error("Failed to navigate to Job Search")
                    return
                
                # Process each Service Order
                for i, so_number in enumerate(so_numbers, 1):
                    self.logger.info(f"\n{'='*50}")
                    self.logger.info(f"Processing {i}/{len(so_numbers)}: {so_number}")
                    self.logger.info(f"{'='*50}")
                    
                    try:
                        success = await self.process_service_order(page, so_number)
                        
                        if success:
                            self.stats["success"] += 1
                            self.logger.info(f"✅ Successfully processed {so_number}")
                        else:
                            self.stats["failed"] += 1
                            self.logger.error(f"❌ Failed to process {so_number}")
                            # Log failed SO
                            with open(self.download_dir / "failed_orders.txt", "a") as f:
                                f.write(f"{so_number}\n")
                        
                    except Exception as e:
                        self.logger.error(f"Critical error processing {so_number}: {e}")
                        self.stats["failed"] += 1
                        await page.screenshot(path=f"critical_error_{so_number}.png")
                    
                    # Return to search page for next iteration (if not last)
                    if i < len(so_numbers):
                        await self.return_to_search(page)
                
                # Print summary
                self.print_summary()
                
            except Exception as e:
                self.logger.error(f"Automation failed: {e}")
                await page.screenshot(path="automation_error.png")
                
            finally:
                # Cleanup
                try:
                    await browser.close()
                except:
                    pass
                
                try:
                    await playwright.stop()
                except:
                    pass
                
        except Exception as e:
            self.logger.error(f"Setup failed: {e}")
        
        finally:
            # Always set end time
            self.stats["end_time"] = datetime.now()
            self.print_summary()
    
    def print_summary(self):
        """Print automation summary"""
        duration = self.stats["end_time"] - self.stats["start_time"]
        
        self.logger.info("\n" + "="*60)
        self.logger.info("AUTOMATION SUMMARY")
        self.logger.info("="*60)
        self.logger.info(f"Total Service Orders: {self.stats['total']}")
        self.logger.info(f"Successfully processed: {self.stats['success']}")
        self.logger.info(f"Failed: {self.stats['failed']}")
        self.logger.info(f"Skipped: {self.stats['skipped']}")
        self.logger.info(f"Duration: {duration}")
        self.logger.info(f"Download location: {self.download_dir.absolute()}")
        
        if self.stats['total'] > 0:
            success_rate = (self.stats['success'] / self.stats['total']) * 100
            self.logger.info(f"Success Rate: {success_rate:.1f}%")
        
        self.logger.info("="*60)
        
        # Save summary to file
        summary_file = self.download_dir / "automation_summary.txt"
        with open(summary_file, 'w') as f:
            f.write("Automation Summary\n")
            f.write("="*40 + "\n")
            f.write(f"Start Time: {self.stats['start_time']}\n")
            f.write(f"End Time: {self.stats['end_time']}\n")
            f.write(f"Duration: {duration}\n")
            f.write(f"Total Orders: {self.stats['total']}\n")
            f.write(f"Success: {self.stats['success']}\n")
            f.write(f"Failed: {self.stats['failed']}\n")
            if self.stats['total'] > 0:
                f.write(f"Success Rate: {(self.stats['success']/self.stats['total']*100):.1f}%\n")
    
    def run(self):
        """Run the automation (synchronous wrapper)"""
        self.logger.info("Starting CRM Automation")
        
        try:
            # Install Playwright browsers if needed
            self.logger.info("Checking Playwright installation...")
            os.system("playwright install chromium")
            
            # Run async automation
            asyncio.run(self.run_automation())
            
        except KeyboardInterrupt:
            self.logger.info("Automation interrupted by user")
        except Exception as e:
            self.logger.error(f"Automation failed: {e}")
        finally:
            self.logger.info("Automation completed")

if __name__ == "__main__":
    # Check if config file exists
    if not Path("config.json").exists():
        print("Error: config.json not found!")
        print("Please run app.py first to create configuration.")
        sys.exit(1)
    
    # Run automation
    automation = CRMAutomation()
    automation.run()