# INSTRUCTIONS

- Clone repo
- npm install
- run the `dev-server` script
- Use the following command to side load into Excel.
  - You will need to install the npm package office-addin-dev-settings globally
  - You will need to replace the path with the full path to the directory you cloned to
    - `office-addin-dev-settings sideload  --document "<path>\doc\testfn-test.xlsx" --app Excel <path>manifest.xml" Desktop`
- The workbook should calculate straight away, if it doesn't, make a change to Sheet1!A1 to another number and that change will flow to all cells causing a full recalculation
