from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from random import randint

import traceback
import os
import shutil
import json
import time
import multiprocessing as mp


class ReportScrapper(object):
    """
    Initializing the Report Scrapper
    """
    def __init__(self):
        try:
            # Parse the configuration
            with open('config.json') as json_file:
                config = json.load(json_file)
            self.general = config["general"]
            self.browserConfig = config["browser"]
            self.credentials = config["credentials"]
            self.downloadLocation = config["storage"]
            self.search = config["searchConfig"]
        except Exception as ex:
            print ("Exception in the configuration file")
            print (ex)
            exit(1)

    """
    SSO Credentials Parser
    """
    def sso(self, driver):
        # Wait for the Osiris Home page to be loaded
        try:
            WebDriverWait(driver, self.general["ssoWaitTime"]).until(
                ec.title_is("USC Shibboleth Single Sign-on")
            )
        except TimeoutException:
            print ("SSO page title mismatch")
        except Exception as ex:
            print ("SSO page encountered unusual exception")
            print (ex)
            print ("Will start with a new browser")
            return False

        # Match the title of the SSO page
        if driver.title == "USC Shibboleth Single Sign-on":
            try:
                usc_id = driver.find_element_by_id("username")
                usc_id.clear()
                usc_id.send_keys(self.credentials["username"])

                pwd = driver.find_element_by_id("password")
                pwd.clear()
                pwd.send_keys(self.credentials["password"])

                # Sign-in button
                driver.find_element_by_name("_eventId_proceed").click()
                return True
            except Exception as ex:
                print ("SSO module Error")
                print (ex)
                return False
        else:
            print ("Couldn't locate the SSO page. Restarting.")
            return False

    """
    Fetch the Web driver
    """
    def get_driver(self, download=True, temp_args=None):
        temp_path = ""

        if "name" in self.browserConfig and self.browserConfig['name'] == 'chrome':
            chrome_options = webdriver.ChromeOptions()
            if download:
                while True:
                    random_no = randint(0, 100)
                    file_creation = False
                    if not os.path.isdir(
                            os.path.join(self.downloadLocation["path"], temp_args, "test" + str(random_no))):
                        temp_path = os.path.join(self.downloadLocation["path"], temp_args, "test" + str(random_no))
                        try:
                            os.makedirs(temp_path)
                            file_creation = True
                        except Exception as ex:
                            print ("Exception encountered in creating temporary folder.Retry with another folder.")
                            print ex
                        if file_creation:
                            break

                chrome_options.add_experimental_option('prefs', {
                    "plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}],
                    "download": {
                        "prompt_for_download": False,
                        "default_directory": temp_path
                    }
                })

            if "incognito" in self.browserConfig:
                # Code for incognito mode
                chrome_options.add_argument("--incognito")

            # Load the browser
            driver = webdriver.Chrome(chrome_options=chrome_options)
        else:
            # Code for mozilla
            fp = webdriver.FirefoxProfile()

            if download:
                while True:
                    random_no = randint(0, 100)
                    file_creation = False
                    if not os.path.isdir(
                            os.path.join(self.downloadLocation["path"], temp_args, "test" + str(random_no))):
                        temp_path = os.path.join(self.downloadLocation["path"], temp_args, "test" + str(random_no))
                        try:
                            os.makedirs(temp_path)
                            file_creation = True
                        except Exception as ex:
                            print (
                                "Exception encountered in creating temporary folder.Retry with another folder.")
                            print ex
                        if file_creation:
                            break

                mime_types = "application/pdf,application/vnd.adobe.xfdf,application/vnd.fdf," \
                             "application/vnd.adobe.xdp+xml"
                fp.set_preference("browser.download.folderList", 2)
                fp.set_preference("browser.download.manager.showWhenStarting", False)
                fp.set_preference("browser.download.dir", temp_path)
                fp.set_preference("browser.helperApps.neverAsk.saveToDisk", mime_types)
                fp.set_preference("plugin.disable_full_page_plugin_for_types", mime_types)
                fp.set_preference("pdfjs.disabled", True)

            driver = webdriver.Firefox(firefox_profile=fp)

        # URL of the site
        driver.get(self.general["url"])

        return driver, temp_path

    def osiris_home(self, driver):

        # Initiate the automation
        if not self.sso(driver):
            if driver.title == "Osiris - Home":
                return True
            return False
        else:
            # Wait for the Osiris Home page to be loaded
            try:
                WebDriverWait(driver, self.general["osirisLoadTime"]).until(
                    ec.title_is("Osiris - Home")
                )
                return True
            except TimeoutException:
                print ("Osiris page didn't load in " + str(self.general["osirisLoadTime"])
                       + " seconds. Might be a Network issue. Asking for a Retry.")
                return False
            except Exception as ex:
                print ("Unknown exception encountered while loading Osiris home page. Title could not be matched.")
                print (ex)
                return False

    def osiris_global_reports_page(self, driver):
        # Set state to that of the global reports page

        # Setup the site to the home page of OSIRIS
        if not self.osiris_home(driver):
            return False

        counter = 1
        while counter < 6:
            # Select the global tab
            try:
                driver.find_element_by_xpath(
                    "//a[@title='Images of annual and interim reports']").click()
                break
            except NoSuchElementException:
                counter += 1
                time.sleep(5)
            except Exception as ex:
                print ("Unexpected exception encountered. Recommend restarting.")
                print (ex)
                counter = 6

        if counter == 6:
            print ("Couldn't locate the Global Tab. Will Retry with new window")
            return False

        counter = 1
        while counter < 4:
            # Confirmation of the Global page
            active_tab = driver.find_element_by_class_name("mainTabLinkOn")
            if active_tab.text == "Global Reports":
                break
            else:
                counter += 1
                time.sleep(2)

        if counter == 4:
            print ("Couldn't locate the Global Tab. Will Retry with new window")
            return False

        return True

    def populate_search_constraints(self, driver, country, year, company):
        try:
            if "words" in self.search and len(self.search["words"]) > 0:
                driver.find_element_by_id("ContentContainer1_ctl00_Content_WordsTextBox").clear()
                driver.find_element_by_id("ContentContainer1_ctl00_Content_WordsTextBox").send_keys(
                    self.search["words"])

            if len(company) > 0:
                driver.find_element_by_id("ContentContainer1_ctl00_Content_CompanyTextBox").clear()
                driver.find_element_by_id("ContentContainer1_ctl00_Content_CompanyTextBox").send_keys(
                    company)

            Select(driver.find_element_by_id("ContentContainer1_ctl00_Content_CountriesDropDownList")).select_by_value(
                country)

            Select(driver.find_element_by_id("ContentContainer1_ctl00_Content_ReportYearList")).select_by_value(
                year)

            if "type" in self.search and len(self.search["type"]) > 0:
                Select(driver.find_element_by_id("ContentContainer1_ctl00_Content_ReportTypeList")).select_by_value(
                    self.search["type"])

            driver.find_element_by_id("ContentContainer1_ctl00_Content_SearchButton").click()
        except Exception as ex:
            print ("Search Module Error")
            print (ex)
            print traceback.format_exc()
            print ("Retrying for ", country, " for year ", str(year))
            return False

        # Wait for the Osiris Search page to be load results
        try:
            # Delay in ajax
            WebDriverWait(driver, self.general["searchLoadTime"]).until(
                ec.invisibility_of_element_located(
                    (By.ID, "ContentContainer1_ctl00_Content_SearchLabel"))
            )
        except TimeoutException:
            print ("Search results didn't load within " + str(self.general["osirisLoadTime"])
                   + " seconds. Search label still visible. WIll retry entire search.")
            return False
        except Exception as ex:
            print ("Search results process didn't go as per procedure")
            print (ex)
            return False
        return True

    @staticmethod
    def delete_folder(path):
        counter = 1
        while True:
            try:
                shutil.rmtree(path)
                break
            except Exception as ex:
                print ("Unable to delete the temporary directory : " + path + " . Attempt : " + str(counter))
                print ex
                print ("Will retry next deletion attempt in few seconds.")
                counter += 1
                if counter > 4:
                    print ("Unable to delete temporary directory within 3 tries. Please manually delete folder : "
                           + path)
                    break

    @staticmethod
    def empty_folder(path):
        for file1 in os.listdir(path):
            file_path = os.path.join(path, file1)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as ex:
                print ("Unable to empty the temporary directory")
                print ex

    def enter_file_in_log(self, sno, name, date, size, version, page_number, temp_args, year):
        content = str(sno) + "," + name + "," + str(date) + "," + str(version) + "," + str(size) + "\n"
        page_number = str(page_number)
        year = str(year)

        if not os.path.isfile(
                os.path.join(self.downloadLocation["path"], temp_args, year + "_" + page_number + ".txt")):
            try:
                print ("Create a new log file for this page number : " + page_number)
                file1 = open(os.path.join(self.downloadLocation["path"], temp_args, year + "_" + page_number + ".txt"),
                             "w+")
                file1.write("SNo,Name,Closing Date,Version,Size (kilobytes)\n")
                file1.close()
            except Exception as ex:
                print ("Exception encountered in creation of the log tracker file. Headings not inserted.")
                print ex

        try:
            file1 = open(os.path.join(self.downloadLocation["path"], temp_args, year + "_" + page_number + ".txt"),
                         "a+")
            # write the data to the files
            file1.write(content)
            # Close the file
            file1.close()
        except Exception as ex:
            print ("Exception encountered in data insertion in the lof tracker file. Content is being skipped")
            print ex
            print ("Content being skipped is : ", content)

    def row_expansion(self, driver, temp_download, row_number, columns, expansion=True):
        # If the expansion of the row is possible
        try:
            # Ask the browser to expand that row
            columns[0].find_element_by_tag_name("input").click()
        except Exception:
            # No expansion icon present, so proceed normally
            return False, driver

        # Wait for the expansion to take place
        expansion_counter = 0
        while expansion_counter < 4:
            time.sleep(5)
            try:
                # New Result table present
                result_table = driver.find_element_by_id(
                    "ContentContainer1_ctl00_Content_DataGridReportList")
                columns = result_table.find_elements_by_tag_name("tr")[
                    row_number].find_elements_by_tag_name("td")
                image_url = columns[0].find_element_by_tag_name("input").get_attribute("src")

                if expansion:
                    image = "Min.gif"
                else:
                    # ignoring P of plus
                    image = "lus.gif"

                # The row was expanded
                if image_url[len(image_url) - 7:] == image:
                    break
                else:
                    # Lets give it some more time
                    expansion_counter += 1
            except Exception:
                # Some exception was encountered, so lets give it some more time
                expansion_counter += 1

        if expansion_counter == 4:
            # Restart the browser, if something breaks it will refer the main try catch
            print ("Will restart the browser as the expansion part encountered some issue")
            driver.quit()
            self.delete_folder(temp_download)
            return False, None

        # Expansion was successful
        return True, driver

    '''
    Function that initiates the download of the files of a particular page
    '''

    def initiate_download(self, page_number, country, year, company, output):
        row_number = 1
        actual_file_count = 1
        expand_present = False
        expansion_sno = None
        expansion_delta = 0

        # Prepare the temp_args for the file download location
        if len(country) == 0:
            temp_args = "allCountries"
        else:
            temp_args = country

        if len(company) > 0:
            temp_args = os.path.join(temp_args, company)

        print ("Initiating download for the page number " + str(page_number) + " for " + country + " for " + str(year))

        while True:
            driver, temp_download = self.get_driver(True, temp_args)

            if not self.osiris_global_reports_page(driver):
                driver.quit()
                self.delete_folder(temp_download)
                print (
                "Re-Initiating download for the page number " + str(page_number) + " for " + country + " for " + str(
                    year))
                print ("Number of files with their all versions downloaded from this page or the row number : " + str(
                    row_number - 1))
                continue

            if not self.populate_search_constraints(driver, country, year, company):
                driver.quit()
                self.delete_folder(temp_download)
                print (
                "Re-Initiating download for the page number " + str(page_number) + " for " + country + " for " + str(
                    year))
                print ("Number of files with their all versions downloaded from this page or the row number : " + str(
                    row_number - 1))
                continue

            if page_number > 1:
                try:
                    # Update the page number
                    driver.find_element_by_id(
                        "ContentContainer1_ctl00_Content_CurrentPage").\
                        send_keys(Keys.BACKSPACE, Keys.BACKSPACE, Keys.BACKSPACE,
                                  Keys.BACKSPACE, page_number, Keys.RETURN)
                except Exception:
                    print ("Couldn't find the page input field box. Will restart with new window.")
                    driver.quit()
                    self.delete_folder(temp_download)
                    print ("Re-Initiating download for the page number " + str(page_number)
                           + " for " + country + " for " + str(year))
                    print ("Number of files with their all versions downloaded from this page or the row number : "
                           + str(row_number - 1))
                    continue

                try:
                    # Delay in ajax
                    WebDriverWait(driver, self.general["searchLoadTime"]).until(
                        ec.text_to_be_present_in_element_value(
                            (By.ID, "ContentContainer1_ctl00_Content_CurrentPage"), str(page_number))
                    )
                except Exception as ex:
                    print ("Pagination failed to perform its job for page number " + str(page_number) + " for "
                           + country + " for " + str(year) + " while for page initial state. Will retry.")
                    print (ex)
                    driver.quit()
                    self.delete_folder(temp_download)
                    print ("Re-Initiating download for the page number " + str(page_number))
                    print ("Number of files with their all versions downloaded from this page or the row number : "
                           + str(row_number - 1))
                    continue

            # Check if the results were published
            try:
                # Result table present
                result_table = driver.find_element_by_id(
                    "ContentContainer1_ctl00_Content_DataGridReportList")

                # Total number of rows in the result table
                total_rows = len(result_table.find_elements_by_tag_name("tr"))

                # Check if there was a restart in the browser due to an error for the expansion case
                if row_number < actual_file_count:
                    expand_present, driver = \
                        self.row_expansion(driver, temp_download,
                                           row_number, result_table.find_elements_by_tag_name("tr")[
                                               row_number].find_elements_by_tag_name("td"))

                    # Row expansion failed
                    if driver is None:
                        continue

                    # If expansion was triggered, then update the element references
                    if expand_present:
                        try:
                            # Result table present
                            result_table = driver.find_element_by_id(
                                "ContentContainer1_ctl00_Content_DataGridReportList")
                            # Storing the Expansion Sno
                            expansion_sno = result_table.find_elements_by_tag_name("tr")[
                                row_number].find_elements_by_tag_name("td")[1].text
                        except Exception:
                            # Some error occurred and couldn't update the references so quit and start again
                            print ("Will restart the browser as it failed to update the"
                                   + " table reference after the expansion of the rows")
                            driver.quit()
                            self.delete_folder(temp_download)
                            driver = None
                            continue

                # If it is a new iteration then calculate the current page value
                try:
                    current_page = int(driver.find_element_by_id(
                        "ContentContainer1_ctl00_Content_CurrentPage").get_attribute("value"))
                except Exception:
                    current_page = 1

                while row_number < total_rows and driver is not None:
                    # Check if some of the data moved to the some page and we are not on that page
                    if actual_file_count / 26 > expansion_delta or current_page < page_number + expansion_delta:
                        # Check if we just crossed the page limit
                        if actual_file_count / 26 > expansion_delta:
                            expansion_delta += 1

                        print ("The data moved to the some page. Going to that page for the data "
                               + "supposed to be on the page " + str(page_number) + ".")
                        try:
                            # Send the new page number
                            driver.find_element_by_id(
                                "ContentContainer1_ctl00_Content_CurrentPage").\
                                send_keys(
                                Keys.BACKSPACE, Keys.BACKSPACE,
                                Keys.BACKSPACE, Keys.BACKSPACE, (page_number + expansion_delta), Keys.RETURN)

                            # Update the current page number
                            current_page = page_number + expansion_delta
                        except Exception:
                            print ("Couldn't find the page input field box. Will restart with new window. "
                                   + "This for data shifted to next page due to expansion.")
                            driver.quit()
                            self.delete_folder(temp_download)
                            print ("Re-Initiating download for the page number " + str(
                                page_number) + " for " + country + " for " + str(year))
                            print ("Number of files with their all versions downloaded from this page "
                                   + "or the row number : " + str(row_number - 1))
                            driver = None
                            continue

                        try:
                            # Delay in ajax
                            WebDriverWait(driver, self.general["searchLoadTime"]).until(
                                ec.text_to_be_present_in_element_value(
                                    (By.ID, "ContentContainer1_ctl00_Content_CurrentPage"),
                                    str(page_number + expansion_delta))
                            )
                        except Exception as ex:
                            print ("Pagination failed to perform its job for page number " + str(
                                page_number) + " for " + country + " for " + str(year) +
                                   " as the data moved to its second page. Will retry.")
                            print ex
                            driver.quit()
                            self.delete_folder(temp_download)
                            print ("Re-Initiating download for the page number " + str(page_number))
                            print ("Number of rows parsed for this page : " + str(row_number - 1))
                            driver = None
                            continue

                        try:
                            # Result table present
                            result_table = driver.find_element_by_id(
                                "ContentContainer1_ctl00_Content_DataGridReportList")
                        except Exception:
                            print ("Failed to update the result table reference upon jump to page as data moved.")
                            driver.quit()
                            self.delete_folder(temp_download)
                            print ("Re-Initiating download for the page number " + str(page_number))
                            print ("Number of rows parsed for this page : " + str(row_number - 1))
                            driver = None
                            continue

                    # Get the columns list but before that check if the data is on the next page
                    if expansion_delta > 0:
                        # Calculate which row to pick but first check for the expansions
                        if len(result_table.find_elements_by_tag_name("tr")) == actual_file_count - 26 * expansion_delta:
                            row_number += 1
                            actual_file_count = row_number
                            continue

                        columns = result_table.find_elements_by_tag_name("tr")[
                            actual_file_count - 26 * expansion_delta]. \
                            find_elements_by_tag_name("td")
                    else:
                        # Normal loop but check for any expansions
                        if expand_present and len(result_table.find_elements_by_tag_name("tr")) == actual_file_count:
                            row_number += 1
                            actual_file_count = row_number
                            continue

                        columns = result_table.find_elements_by_tag_name("tr")[
                            actual_file_count].find_elements_by_tag_name(
                            "td")

                    '''
                    If the next row says anything about expansion
                    '''
                    if len(columns[1].text) > 1:
                        # If expansion already present and compression possible when on same page
                        if expand_present and actual_file_count < total_rows:
                            # Unset the flag and Sno
                            expand_present = False
                            expansion_sno = None

                            # Call for row compression
                            compress, driver = self.row_expansion(driver, temp_download, row_number,
                                                                  result_table.find_elements_by_tag_name("tr")[
                                                                      row_number].find_elements_by_tag_name("td"),
                                                                  False)

                            # Update the counters
                            row_number += 1
                            actual_file_count = row_number

                            # Row contraction failed
                            if driver is None:
                                # Since a new window will be initiated so forget the expansion related problems
                                break

                            # If compression was triggered, then update the element references
                            if compress:
                                try:
                                    # Result table present
                                    result_table = driver.find_element_by_id(
                                        "ContentContainer1_ctl00_Content_DataGridReportList")
                                    # Columns list - used row_number to mark the beginning of the inner loop
                                    columns = result_table.find_elements_by_tag_name("tr")[
                                        row_number].find_elements_by_tag_name("td")
                                except Exception as ex:
                                    # Some error occurred and couldn't update the references so quit and start again
                                    driver.quit()
                                    self.delete_folder(temp_download)
                                    driver = None
                                    print ("Triggering compression failed in updation new result table generated")
                                    print (ex)
                                    # Since a new window will be initiated so forget the expansion
                                    break

                        # Restart the browser. No trouble of going back to the previous page and setting the flags
                        elif expansion_delta > 0:
                            driver.quit()
                            self.delete_folder(temp_download)
                            driver = None

                            # Since a new window will be initiated so forget the expansion
                            row_number += 1
                            actual_file_count = row_number
                            expand_present = False
                            expansion_delta = 0
                            print ("Reset browser to previous state. Data on different page, it became complicated.")
                            break

                    # Now process normally since we pass compression and new page checks
                    try:
                        # Check for already an expansion
                        if not expand_present:
                            # See if you can expand the row
                            expand_present, driver = self.row_expansion(driver, temp_download, row_number, columns)

                            # Row expansion failed
                            if driver is None:
                                break

                            # If expansion was triggered, then update the element references
                            if expand_present:
                                try:
                                    # Result table present
                                    result_table = driver.find_element_by_id(
                                        "ContentContainer1_ctl00_Content_DataGridReportList")
                                    # Columns list - used row_number to mark the beginning of the inner loop
                                    columns = result_table.find_elements_by_tag_name("tr")[
                                        row_number].find_elements_by_tag_name("td")
                                    expansion_sno = columns[1].text
                                except Exception:
                                    # Some error occurred and couldn't update the references so quit and start again
                                    print ("Failed to update the table columns after expansion")
                                    driver.quit()
                                    self.delete_folder(temp_download)
                                    driver = None
                                    break

                        # If report language matches the requested language
                        if columns[6].text in self.search["language"]:
                            # Store the main window handler
                            parent_window = driver.current_window_handle

                            # Make sure that the temp folder is empty
                            self.empty_folder(temp_download)

                            # Open the pdf
                            report_link = columns[4].find_element_by_tag_name("a")
                            report_link.click()

                            company_name = columns[2].text.replace(
                                "/", "-")
                            company_name_list = company_name.split(" ")

                            # Check if the company name is larger than 3 words
                            if len(company_name_list) > 3:
                                company_name = company_name[:len(company_name_list[0]) + 1 + len(company_name_list[1])
                                                             + 1 + len(company_name_list[2])]

                            # Prepare the file Name
                            if expand_present:
                                # If this is a part of the expansion
                                new_file_name = expansion_sno \
                                                + "_" + company_name + "_" + report_link.text.replace("/", "-") \
                                                + "_v_" + str(actual_file_count - row_number) + ".pdf"
                            else:
                                # Normal file name preparation
                                new_file_name = str(
                                    columns[1].text) + "_" + company_name + "_" + report_link.text.replace(
                                    "/", "-") + ".pdf"

                            # Check if this file name already exists. Redundant but might be useful
                            index = 0
                            while True:
                                if os.path.isfile(os.path.join(self.downloadLocation["path"],
                                                               temp_args, new_file_name)):
                                    # Prepare the file Name
                                    if expand_present:
                                        # If this is a part of the expansion
                                        new_file_name = expansion_sno \
                                                        + "_" + company_name + "_" \
                                                        + report_link.text.replace("/", "-")\
                                                        + "_v_" + str(actual_file_count - row_number) + "_Copy_" \
                                                        + str(index) + ".pdf"
                                    else:
                                        # Normal file name preparation
                                        new_file_name = str(
                                            columns[1].text) + "_" + company_name + "_" + report_link.text.replace(
                                            "/", "-") + "_Copy_" + str(index) + ".pdf"

                                    index += 1
                                else:
                                    # The file name is unique
                                    break

                            # Delay to switch to the new tab
                            counter = 0
                            while counter < 4:
                                time.sleep(1)
                                if len(driver.window_handles) > 1:
                                    driver.switch_to.window(driver.window_handles[1])
                                    break
                                counter += 1

                            # Fail safe
                            if counter == 4:
                                print ("File download tab didn't open. Restarting for page number " + str(page_number)
                                       + " for " + country + " for " + str(year))
                                driver.quit()
                                self.delete_folder(temp_download)
                                driver = None
                                continue

                            # Check if the download is complete
                            while True:
                                if len(os.listdir(temp_download)) > 0:
                                    filename = os.listdir(temp_download)[0]
                                    while filename.endswith("crdownload"):
                                        time.sleep(10)
                                        filename = os.listdir(temp_download)[0]
                                    break
                                else:
                                    print ("Waiting for download to begin for page " + str(page_number)
                                           + " for file number " + str(row_number) + " for " + country
                                           + " for " + str(year))
                                    if expand_present:
                                        print ("This file has multiple versions. Downloading version : "
                                               + str(actual_file_count - row_number))
                                    time.sleep(5)

                            # Close this new download tab
                            driver.close()

                            # Precaution check for another extra opened tabs
                            if len(driver.window_handles) > 1:
                                for tab in driver.window_handles:
                                    if not tab == parent_window:
                                        driver.switch_to.window(tab)
                                        driver.close()

                            # Switch to the parent tab
                            driver.switch_to.window(parent_window)

                            # Shift the file
                            try:
                                os.rename(os.path.join(temp_download, filename),
                                          os.path.join(self.downloadLocation["path"], temp_args, new_file_name))
                            except Exception as ex:
                                dir_list = os.listdir(temp_download)
                                if len(dir_list) > 0 and dir_list[0] == filename and filename.endswith("pdf"):
                                    # If unable to rename the file but the file exists, rename the directory itself
                                    print ("Unable to rename file, so will try to rename the folder itself")
                                    try:
                                        os.rename(temp_download,
                                                  os.path.join(self.downloadLocation["path"], temp_args, new_file_name))
                                        self.enter_file_in_log(
                                            expansion_sno or str(columns[1].text), company_name, report_link.text,
                                            columns[7].text, actual_file_count - row_number, page_number,
                                            temp_args, year)
                                        print ("Folder renaming successful, so will skip this file on row "
                                               + str(row_number) + " version " + str(actual_file_count - row_number)
                                               + " for page number " + str(page_number))
                                    except Exception as ex:
                                        print ("Failed to rename the folder.")
                                        print (ex)
                                        print ("File present in one of the test folders")
                                    finally:
                                        row_number += 1

                                    print ("Will restart with a new window for page number "
                                           + str(page_number) + " for " + country + " for " + str(year))
                                    driver.quit()
                                    driver = None
                                    break
                                else:
                                    # This was a genuine error
                                    print ("File to rename doesn't exists, restart in new window for page number "
                                           + str(page_number) + " for " + country + " for " + str(year))
                                    driver.quit()
                                    self.delete_folder(temp_download)
                                    driver = None
                                    break

                            # Write to the log file
                            self.enter_file_in_log(expansion_sno or str(
                                columns[1].text), company_name, report_link.text, columns[7].text,
                                actual_file_count - row_number, page_number, temp_args, year)

                    except Exception as ex:
                        print ("Exception Encountered during file download process. Close the browser and restart.")
                        print (ex)
                        print traceback.format_exc()
                        driver.quit()
                        self.delete_folder(temp_download)
                        driver = None
                        print ("Re-Initiating download for the page number " + str(page_number)
                               + " for " + country + " for " + str(year))
                        print ("Number of rows processed for this page : "
                        + str(row_number - 1))
                        break

                    # Increase the row count
                    if driver is not None:
                        # Increase the actual file count
                        actual_file_count += 1

                        # Report count to only increase when its a new report and not a variant
                        if not expand_present:
                            row_number += 1
                            actual_file_count = row_number
                            expansion_delta = 0

                # Exit condition plus if the last part was expanded and resulted in browser closure
                if row_number == total_rows:
                    if driver is not None:
                        driver.quit()
                        self.delete_folder(temp_download)
                    break
            except Exception as ex:
                # Should encounter this part if result page fails to populate
                print ("Result page failed to populate.")
                print (ex)
                print traceback.format_exc()
                driver.quit()
                self.delete_folder(temp_download)
                print ("Re-Initiating download for the page number " + str(page_number)
                       + " for " + country + " for " + str(year))
                print ("Number of files with their all versions downloaded from this page or the row number : "
                       + str(row_number - 1))
                continue

        # Rename the txt logs to csv
        try:
            # There exists an old csv
            if os.path.isfile(
                    os.path.join(self.downloadLocation["path"], temp_args,
                                 str(year) + "_" + str(page_number) + ".csv")):
                os.remove(
                    os.path.join(self.downloadLocation["path"], temp_args, str(year) + "_" + str(page_number) + ".csv"))

            if os.path.isfile(
                    os.path.join(self.downloadLocation["path"], temp_args,
                                 str(year) + "_" + str(page_number) + ".txt")):
                os.rename(
                    os.path.join(self.downloadLocation["path"], temp_args, str(year) + "_" + str(page_number) + ".txt"),
                    os.path.join(self.downloadLocation["path"], temp_args, str(year) + "_" + str(page_number) + ".csv"))
            else:
                print "Couldn't locate the log file to be renamed. By passing this."
        except Exception as ex:
            print "Unknown exception encountered while changing the log file extension"
            print ex
            print "Continuing the normal execution"

        output.put((page_number, "Successfully downloaded " + str(row_number - 1) + " files for the page number "
                    + str(page_number) + " for the search result for country " + country
                    + " for year " + str(year)))
        print ("Exiting process for page " + str(page_number))

    def run(self):
        # Loops over the company
        if "company" in self.search and len(self.search["company"]) > 0:
            # Loop in for the company
            for company in self.search["company"]:
                # Loops over the years
                if "year" in self.search and len(self.search["year"]) > 0:
                    # Loop in for the years
                    for year in self.search["year"]:
                        # Cartesian Product
                        if "country" in self.search and len(self.search["country"]) > 0:
                            # Loop in for the years
                            for country in self.search["country"]:
                                page_count = 0
                                max_threads = self.general["maxThreads"]

                                while True:
                                    driver, _ = self.get_driver(False)

                                    if not self.osiris_global_reports_page(driver):
                                        driver.quit()
                                        continue

                                    if not self.populate_search_constraints(driver, country, year, company):
                                        driver.quit()
                                        continue

                                    # Check if the results were published
                                    try:
                                        # Result table present
                                        driver.find_element_by_id(
                                            "ContentContainer1_ctl00_Content_DataGridReportList")

                                    except Exception:
                                        try:
                                            # No matching search results
                                            driver.find_element_by_id("ContentContainer1_ctl00_Content_Label2")
                                            print ("Results not found for " + country + " for the year " + str(year))
                                            driver.quit()
                                            # Get out of the loop for this configuration
                                            break
                                        except Exception as ex:
                                            # Unknown error encountered
                                            print ("Exception encountered in displaying search results for " + country +
                                                   " for the year " + str(year) + ". Will restart analysing.")
                                            print (ex)
                                            driver.quit()
                                            continue

                                    # Capture Page Count
                                    try:
                                        captured_page_count = \
                                            driver.find_element_by_id("ContentContainer1_ctl00_Content_PagesLabel").text
                                        captured_page_count = captured_page_count[3:]

                                        index = 1
                                        while True:
                                            if captured_page_count[index] == ' ':
                                                break
                                            index += 1
                                        page_count = int(captured_page_count[:index])
                                        driver.quit()
                                        break
                                    except Exception:
                                        # If no pagination text box present
                                        try:
                                            # Result table present reconfirm
                                            driver.find_element_by_id(
                                                "ContentContainer1_ctl00_Content_DataGridReportList")
                                            page_count = 1
                                            driver.quit()
                                            break
                                        except Exception as ex:
                                            # Unknown error encountered
                                            print ("Unable to get the pagination count for search results for "
                                                   + country + " for the year " + str(year)
                                                   + ". Will restart analysing.")
                                            print (ex)
                                            print traceback.format_exc()
                                            driver.quit()
                                            continue

                                if page_count > 0:
                                    print ("Total number of pages : " + str(page_count))
                                    # Setup the page starting point
                                    if "specificPageStart" in self.search and self.search["specificPageStart"] > 0:
                                        page_start = self.search["specificPageStart"]
                                    else:
                                        page_start = 1

                                    if "specificPageEnd" in self.search \
                                            and self.search["specificPageEnd"] - page_start <= page_count \
                                            and self.search["specificPageEnd"] > 0:
                                        page_end = self.search["specificPageEnd"]
                                    else:
                                        page_end = page_count

                                    # Output Queue
                                    output = mp.Queue()

                                    while True:
                                        temp_page_end = page_start + max_threads - 1
                                        print ("Pages for this iteration " + str(page_start) + " to "
                                               + str(temp_page_end))

                                        # If it already reached the last page
                                        if temp_page_end > page_end:
                                            temp_page_end = page_end

                                            # Check if all the pages have been parsed
                                            if page_start > temp_page_end:
                                                print ("All the pages have been parsed. Terminating program.")
                                                break

                                        # Setup threads
                                        processes = [mp.Process(target=self.initiate_download,
                                                                args=(page, country, year, company, output))
                                                     for page in range(page_start, temp_page_end + 1)]

                                        # Run processes
                                        for p in processes:
                                            p.start()

                                        # Exit the completed processes
                                        for p in processes:
                                            p.join()

                                        # Get process results from the output queue
                                        results = [output.get() for p in processes]

                                        # Print results from all the processes
                                        print(results)

                                        # Setup for the next processes
                                        page_start = temp_page_end + 1
                                        print ("Next Process Start Page " + str(page_start))


if __name__ == '__main__':
    scrapper = ReportScrapper()
    scrapper.run()
