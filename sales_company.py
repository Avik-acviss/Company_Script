from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

# Configure Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--remote-debugging-port=9222")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

# CSV Header including new fields for Fax and Factory Location.
csv_header = [
    "Company Name", "Address", "Phone", "Fax", "Email", "Website",
    "Product(s) & Services", "Factory Location",
    "FTI Representative 1", "FTI Representative 2", "FTI Representative 3",
    "Industrial Club", "Industrial Rep 1", "Industrial Rep 2", "Industrial Rep 3"
]


# Function to safely extract text using an XPath
def safe_get(xpath):
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except Exception:
        return "Not Available"


# Function to safely extract attribute using an XPath
def safe_href(xpath):
    try:
        return driver.find_element(By.XPATH, xpath).get_attribute('href')
    except Exception:
        return "Not Available"


# Function to extract product(s) & services robustly:
def get_products_services():
    try:
        ps_text = driver.find_element(By.ID, "product_service").text.strip()
    except Exception:
        ps_text = ""
    try:
        ps_table_text = driver.find_element(By.XPATH,
                                            "//span[@id='product_service']/following-sibling::table[1]").text.strip()
    except Exception:
        ps_table_text = ""
    combined = (ps_text + " " + ps_table_text).strip()
    return combined if combined else "Not Available"


# List of letters to loop through (A to Z)
letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# Open CSV file for writing scraped data
with open("scraped_data.csv", mode="w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(csv_header)

    for letter in letters:
        print(f"==== Processing Letter: {letter} ====")
        # Construct URL for the letter results page
        letter_url = f"https://ftimember.off.fti.or.th/_layouts/membersearch/resultEN.aspx?ts=4&texts={letter}"
        driver.get(letter_url)
        time.sleep(2)  # Allow page to load

        while True:
            try:
                # Wait for company links to be present on the current page
                wait.until(EC.presence_of_all_elements_located(
                    (By.XPATH, "//td[@align='left']//a[starts-with(@href, 'MemberDetailEN.aspx')]")
                ))
                companies = driver.find_elements(
                    By.XPATH, "//td[@align='left']//a[starts-with(@href, 'MemberDetailEN.aspx')]"
                )
                company_links = [link.get_attribute('href') for link in companies]
                print(f"Processing {len(company_links)} company links on current page.")

                # Save the current window handle (the listing page)
                main_window = driver.current_window_handle

                # Process each company on this page
                for link in company_links:
                    try:
                        # Open detail page in a new tab
                        driver.execute_script("window.open(arguments[0]);", link)
                        new_window = [handle for handle in driver.window_handles if handle != main_window][0]
                        driver.switch_to.window(new_window)
                        time.sleep(2)  # Wait for detail page to load

                        company_name = safe_get("//span[@id='comp_person_name']")
                        address = safe_get("//span[@id='comp_address']")
                        phone = safe_get("//span[@id='addr_telephone']")
                        fax = safe_get("//span[@id='addr_fax']")
                        email = safe_get("//span[@id='addr_email']//a")
                        website = safe_href("(//td[@class='auto-style9']//a)[1]")
                        products_services = get_products_services()
                        factory_location = safe_get("//span[@id='factory_location']")

                        # FTI Representatives
                        fti_reps = driver.find_elements(By.XPATH, "//table[@id='ContactFTI']//tr[2]/td")
                        fti_rep1 = fti_reps[0].text.strip() if len(fti_reps) > 0 else "Not Available"
                        fti_rep2 = fti_reps[1].text.strip() if len(fti_reps) > 1 else "Not Available"
                        fti_rep3 = fti_reps[2].text.strip() if len(fti_reps) > 2 else "Not Available"

                        # Industrial Club and Representatives
                        industrial_sections = driver.find_elements(
                            By.XPATH, "//table[@id='ContactNonFTI']//td[contains(@style, 'background-color')]"
                        )
                        if industrial_sections:
                            # If multiple industrial sections exist, write a row for each.
                            for section in industrial_sections:
                                bolds = section.find_elements(By.XPATH, ".//b")
                                club_name = bolds[0].text.strip() if bolds else "Unknown Club"
                                reps = section.find_elements(By.XPATH, ".//tr[2]/td")
                                rep1 = reps[0].text.strip() if len(reps) > 0 else "Not Available"
                                rep2 = reps[1].text.strip() if len(reps) > 1 else "Not Available"
                                rep3 = reps[2].text.strip() if len(reps) > 2 else "Not Available"
                                writer.writerow([
                                    company_name, address, phone, fax, email, website,
                                    products_services, factory_location,
                                    fti_rep1, fti_rep2, fti_rep3,
                                    club_name, rep1, rep2, rep3
                                ])
                        else:
                            writer.writerow([
                                company_name, address, phone, fax, email, website,
                                products_services, factory_location,
                                fti_rep1, fti_rep2, fti_rep3,
                                "Not Available", "Not Available", "Not Available", "Not Available"
                            ])
                    except Exception as e:
                        print(f"Error processing {link}: {e}")
                    finally:
                        # Close the detail tab and switch back to the listing page
                        driver.close()
                        driver.switch_to.window(main_window)

                # Wait for the pagination section to be present
                wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//tr[contains(@style, 'background-color:#284775')]")
                ))
                # Extract current page number
                try:
                    # This XPath selects the cell that is NOT a link (current page cell)
                    current_page_element = driver.find_element(
                        By.XPATH,
                        "//tr[contains(@style, 'background-color:#284775')]/td/table/tbody/tr/td[not(descendant::a)]"
                    )
                    current_page = int(current_page_element.text.strip())
                    print(f"Current Page: {current_page}")
                except Exception as e:
                    print(f"❌ Unable to determine current page: {e}")
                    break

                # Construct XPath for the next page link based on the current page number
                next_page_xpath = (
                    f"//tr[contains(@style, 'background-color:#284775')]/td/table/tbody/tr/td/"
                    f"a[contains(@href, \"Page${current_page + 1}\")]"
                )
                try:
                    next_page_element = driver.find_element(By.XPATH, next_page_xpath)
                    postback_script = next_page_element.get_attribute("href")
                    print(f"Executing pagination script: {postback_script}")
                    if postback_script.startswith("javascript:"):
                        driver.execute_script(postback_script)
                        time.sleep(3)  # Wait for the new page to load
                    else:
                        print("No valid pagination script found.")
                        break
                except Exception as e:
                    print(f"Next page link not found; reached last page or pagination failed: {e}")
                    break

            except Exception as e:
                print(f"General error on current page for letter {letter}: {e}")
                break

driver.quit()
