# stormanAudit
Scrapes data from SharePoint's Storage Management screen, capturing the data in a list

# basic usage
1. Visit the storage management page at [http://{TOP-LEVEL_SITE_URL}/_layouts/15/storman.aspx](http://{TOP-LEVEL_SITE_URL}/_layouts/15/storman.aspx) (this will get SharePoint to index your site's contents)
2. Wait a few minutes, then refresh the page
3. Run the code in stormanAudit.js
4. Code will create a list matching the name at `stormanAudit.recordsListName` to captures its results
5. Run the code to scrape objects that meet or exceed the number of bytes specified at `stormanAudit.logObjectSizeLowerThreshold`, capturing each as a list item
6. Use a tool such as MS Access to analyze the results in the records list

## details of the code
### Briefly, the code does the following:
1. It checks if the records list (matching the name at `stormanAudit.recordsListName`) exists, and if not creates it + adds fields to capture the details of each site object
2. It scrapes the first page of site objects from the storage management page
   * It captures each site object on the page in `stormanAudit.arrEntries`
   * For each object that meets the threshold at `stormanAudit.logObjectSizeLowerThreshold`, we check if it's a link that lets us drill down to its sub-objects
     * If an object has sub-objects, we append an element to `stormanAudit.arrSubentries` to queue the retrieval of those
   * If the last site object shown on the page meets the threshold at `stormanAudit.logObjectSizeLowerThreshold`, we append an element to `stormanAudit.arrSubentries` to queue the retrieval of the next page of results
3. It loops through the contents of `stormanAudit.arrEntries`, capturing each element as a list item in the records list, appending the promise for each ansynchronous AJAX call to create the item to `stormanAudit.arrCreateItemPromises`
4. It waits for all promises in `stormanAudit.arrCreateItemPromises` to be finished
5. It recursively retrieves the sub-objects and "next pages" currently under `stormanAudit.arrSubentries`, appending the promise to GET each page of results to `stormanAudit.arrPromises`
6. It waits for all promises under `stormanAudit.arrPromises` to be finished
7. It repeatedly executes resursively until we get through the full contents of the site that meet our threshold (detected when `stormanAudit.arrPromises.length <= 0 && stormanAudit.arrEntries.length <= 0`)
