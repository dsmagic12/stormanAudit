# stormanAudit
Scrapes data from SharePoint's Storage Management screen, capturing the data in a list

## basic usage
1. Visit the storage management page at http://<TOP-LEVEL SITE URL>/_layouts/15/storman.aspx (this will get SharePoint to index your site's contents)
2. Wait a few minutes, then refresh the page
3. Code will create a list matching the name at stormanAudit.recordsListName to captures its results
4. Run the code to scrape objects that meet or exceed the number of bytes specified at stormanAudit.logObjectSizeLowerThreshold, capturing each as a list item
5. Use a tool such as MS Access to analyze the results in the records list
