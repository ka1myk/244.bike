/**
 * Configuration for Gmail2GDrive
 * See https://github.com/ahochsteger/gmail2gdrive/blob/master/README.md for a config reference
 */
function getGmail2GDriveConfig() {
  return {
    // Global filter
    "globalFilter": "has:attachment",
    // Gmail label for processed threads (will be created, if not existing):
    "processedLabel": "to-gdrive/processed",
    // Sleep time in milli seconds between processed messages:
    "sleepTime": 100,
    // Maximum script runtime in seconds (google scripts will be killed after 5 minutes):
    "maxRuntime": 280,
    // Only process message newer than (leave empty for no restriction; use d, m and y for day, month and year):
    "newerThan": "3d",
    // Timezone for date/time operations:
    "timezone": "GMT+3",
    // Processing rules:
    "rules": [
      { // Store all pdf attachments from example2@example.com to the folder "Examples/example2"
        "filter": "from:gordeev.d@forwardvelo.ru",
        "folder": "'fromForwardToCsv'",
        "filenameFromRegexp": ".*\.xls$" 
      }
    ]
  };
}
