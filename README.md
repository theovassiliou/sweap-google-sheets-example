# sweap-google-sheets-example

An example of using the Sweap API in Google Sheets.

This is a community-maintained example of how the [Sweap](https://sweap.io) API can be used from within a Google Spreadsheet. This example
should serve as inspiration on what can be done with the API, or public APIs in general.

**Disclaimer**: Me or the Sweap-team can not provide support for this example project. Please keep this in mind when deciding whether to use it.

Instructions for getting started are below. If you have any questions or comments, please create a GitHub issue to let us know. We'd be happy to hear your feedback.

## Requirements

### Sweap Authentication Key

If you integrate this code into your google spreadsheet, before being able to use it you will need Sweap API authentication credentials.

### Google Account

To use Google Sheets, you will also need a Google account. It is up to your discretion to comply with all applicable terms of service when using a service like Google Sheets.

## API Consumption Disclaimer

**Important note:** There is an issue with the add-on, were reopening an existing Sheet that contains Sweap API add-on scripts will evaluate all cells, and this evaluation will count against your API consumption if any.

## Setup

1. Open or create a Google Sheet
2. From the "Extensions" menu select "Apps Script". This opens the Apps Script editor.
3. Create a script file name for example SweapAPI.gs
4. Copy the contents of the SweapAPI.gs file of the repo into it.
5. Modify the script to include your Sweap API credentials.
6. You can close the Apps Script editor or just go back to your Sheet.
7. Use the `ListEvents([eventId])` and `ListGuests(eventName)` functions in your cells. See Usage on how to use them

## Usage

This example offers two functions:

- `ListEvents`
- `ListGuests`

Each function has documentation, that is displayed when you enter the function name in a cell.

Please note that with this example you can not create events, guests or the like. With the provided functions you can only read event or guest information.´

Some usage examples

```gscript
=ListEvents()
    'EventName'         'Start Dae'     'End Date'    'State'
    "My workshop"       "26.07.2022"    "26.07.2022"    DRAFT
    "Birthday Party"    "16.07.2022"    "17.07.2022"    ACTIVE
    "Demo Event"        "13.05.2022"    "14.05.2022"    ACTIVE
    (Or whatever events are available to you, the header is not shown)

=ListEvents("Demo Event)
    'EventName'         'Start Dae'     'End Date'    'State'
    "Demo Event"        "13.05.2022"    "14.05.2022"    ACTIVE
    (If the event is available, the header is not shown)

=ListGuests("Demo Event")
    'LastName'  'FirstName' 'Invitation State'  'Entourage Count'   'Attendance State' 
    "Doe"       "John"      "NONE"              "1"                 "NONE"
    "Vader"     "Darth"     "ACCEPTED"          "0"                 "NONE"
    "the Pooh"  "Winnie"    "ACCEPTED"          "3"                 "NONE"
    (All the guests invited)

=COUNT(ListGuests("Demo Event"))
    "4"
    (Number of guests)

=COUNTIF(ListGuests("Demo Event", "ACCEPTED"))
    (Number of guests with Invitation State ACCEPTED)
```

## Project Status

There is currently no major version released. Therefore, minor version releases may include backward incompatible changes.

See CHANGELOG.md or Releases for more information about the changes.

## Original Documentation

The official documentation of the REST API can be found at <https://documenter.getpostman.com/view/5746225/UVBzmp2n#defe4351-fb75-4d87-a73f-b9e8f0d843c2>

## Inspiration

This example has been inspired by the awesome [google-sheets-example](https://github.com/DeepLcom/google-sheets-example) by DeepL. Thanks a lot, folks.
We have adopted the example by using OAuth2 instead of Basic Authentication.
