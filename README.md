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
6. It waits for all promises under `stormanAudit.arrPromises` to be finished
7. It begins its semi-recursive* functionality by calling `stormanAudit.captureResultsThenGetSubEntries` which does the following:
   * Checks if `stormanAudit.arrPromises.length` is > 0, and if so...
     * Calls `stormanAudit.waitForAllData` to check the promises under `stormanAudit.arrPromises` to see if they are done, removing the completed promises as it goes until the array is empty
     * Calls `stormanAudit.captureFilteredResults` to create list items in the records list for each item scraped from the storage management page(s) into the array `stormanAudit.arrEntries`, capturing each element as a list item in the records list, appending the promise for each ansynchronous AJAX call to create the item to `stormanAudit.arrCreateItemPromises`
     * Note that after looping through `stormanAudit.arrEntries`, the `stormanAudit.captureFilteredResults` function calls `stormanAudit.waitForAllItemsCreated` to wait for all of the promises under `stormanAudit.arrCreateItemPromises` to finish before proceeding
     * Calls `stormanAudit.loopPendingSubEntries` to loop through the queued sub-entry & "next" pages under `stormanAudit.arrSubentries`, removing array elements as it proceeds to request the storage management pages containing that data
   * Checks if `stormanAudit.arrPromises.length` is 0 -and- `stormanAudit.arrEntries.length` is > 0, and if so...
     * Calls `stormanAudit.captureFilteredResults` to create list items in the records list for each item scraped from the storage management page(s) into the array `stormanAudit.arrEntries`, capturing each element as a list item in the records list, appending the promise for each ansynchronous AJAX call to create the item to `stormanAudit.arrCreateItemPromises`
     * Note that after looping through `stormanAudit.arrEntries`, the `stormanAudit.captureFilteredResults` function calls `stormanAudit.waitForAllItemsCreated` to wait for all of the promises under `stormanAudit.arrCreateItemPromises` to finish before proceeding
     * Calls `stormanAudit.loopPendingSubEntries` to loop through the queued sub-entry & "next" pages under `stormanAudit.arrSubentries`, removing array elements as it proceeds to request the storage management pages containing that data
   * Checks if `stormanAudit.arrPromises.length` is 0 -and- `stormanAudit.arrEntries.length` is 0, and if so...
     * Shows a SP.UI.Status message saying the process is done
   * semi-recusive* by this, I'm referring to the back-and-forth 'conversation' that occurs between `stormanAudit.loopPendingSubEntries` and `stormanAudit.captureResultsThenGetSubEntries`. 
     * The `stormanAudit.loopPendingSubEntries` function calls `stormanAudit.captureResultsThenGetSubEntries` when it is finished making the requests to retrieve the pages for the sub-objects and "next" pages that are queued under `stormanAudit.arrSubentries`
     * The `stormanAudit.captureResultsThenGetSubEntries` function calls `stormanAudit.loopPendingSubEntries` whenever it is not done capturing all of the site objects that meet the threshold at `stormanAudit.logObjectSizeLowerThreshold`
