# Diffbot API Excel client

## Preface

This client is an Excel/VBA based solution allowing business analysts to easily access all of Diffbot's API's.

The result is one Excel row per submitted URL containing the parsed returned JSON information.

## Installation

There is no real installation needed as the Excel workbook can be used as a stand-alone product. 
Just make sure your Excel security configuration enables running macro's. There are two ways of doing this:

- Office Button --> Excel options --> Tust Center --> Macro Settings --> Enable all Macro's
- Enable macros when prompted while opening the workbook

## Usage

### Diffbot related parameters

It all starts in the "Diffbot_API" worksheet.

The "Diffbot Excel Interface" table allows you to select the API type (Article - Analyze - Product). 
Your personal token should be provided as well in order to grant access to the requested API.

### Result related parameters

The "Overall Parameters" table contains information related to the way results are returned:

- "Result Worksheet" contains the name of the worksheet where the results should be stored. If the worksheet does not exist it will automatically be created;
- "Start Row" refers to the first row number where results should be stored in the "Result Worksheet". If less than 4 the system will automatically reset it to 4, so as to take care of the result worksheet heading;
- "Start Column" refers to the first column number where results should be stored in the "Result worksheet";

### Bulk vs One by One

Last but not least you can via the "Bulk Processing Parameters" table select if you prefer to make separate calls for every URL's contained in the "List of URL's to Analyze" table (One by One processing) or if you prefer an asynchronous ("Bulk Job") processing in order to retrieve all results at once later on.

If a Bulk processing is choosen, you'll need to enter a bulk "Job Name" for future reference.

### Providing the URL's to process

As already mentionned, all URL's to be analyzed should be stored in the "List of URL's to Analyze" table.

### Making it work

Pressing the "Get and Parse JSON" button will either:

- Make the separate API calls in case of a one-by-one processing and return the results in the "Result Worksheet"
- Post a job containing all URL's for asynchronous processing. In that case the system will return in the "Bulk Processing Feedback" table if the posting was succesful or not, together with a result code (processing stage).

To know if a bulk job has completed, or to enquire on its status, just press the "Update Bulk Job Status" button. The system will then return the status for the entered job name.

Once the Bulk Job has completed, the data can be retrieved by pressing the "Retrieve and Parse Bulk JSON" button. The parsed information will then be returned in the "Result worksheet".


### Error handling

In case the API call would fail (e.g. invalid token, time-out, etc.) an error JSON file will be returned. This file will be parsed as well and made available in the Result Worksheet.

In case the expected parsed result differs from the returned one, the JSON_Parsing module provides an error string called "Str_Parse_Error" containing all encountered errors while trying to parse the returned JSON file.



## Discussion
