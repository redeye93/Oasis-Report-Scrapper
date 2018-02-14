# Oasis-Report-Scrapper
A selenium based scrapper that downloads th reports from the Oasis with the necessary credentials.

This script is build for Oasis access from USC Single sign on. You need to enter your USC credentials in the config.json to allow the script to function properly. You can even tweak the wait times in the config file depending upon your network connection. The script enables download of the pdf in parallel to speed up the tasks. It is done as 1 process per page. However if the number of search pages is less than the max number of threads, then number of process invoked is equivalent to the number of the search pages.

You can also configure it for the mozilla by adding the configuration in the browser part.

You need to give absolute path in the storage directory where it can store the files. 

Additional features include download from particular pages of the search result pages, multiple search criterias. In case you want to tweak them, put the exact names as given in the site.

The program creates temporary directory where it stores the reports for a particular instance that are in transit and also creates a log csv file for manual cross verification of the files downloaded. It also downloads all the versions of the reports published by the comapnies in that fiscal year even if they are under same category, different languages (provided no constraint in the language field in config) etc.

Incase of a failure within the program, it will start from the point it left or simply will resume from last successful download point.