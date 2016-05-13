# Google-Analytics-Manager
Tool to manage your Google Analytics accounts via a Google Sheet.

## Development Backlog (Status)

### General
- [x] Create sheet
- [x] Retrieve API data
- [ ] Functional Documentation
- [ ] Code Documentation

### GA Specific assets
- [x] List Properties
- [x] Add dataValidation (regex) for the required Properties columns
- [ ] Insert / Update Properties
- [x] List Views
- [ ] Add dataValidation (regex) for the required Views columns
- [ ] Insert / Update Views
- [x] List Filter Links
- [ ] Add dataValidation (regex) for the required Filter Links columns
- [ ] Insert / Update Filter Links
- [x] List Custom Dimensions
- [ ] Add dataValidation (regex) for the required Custom Dimensions columns
- [ ] Insert / Update Custom Dimensions
- [ ] List Custom Metrics
- [ ] Add dataValidation (regex) for the required Custom Metrics columns
- [ ] Insert / Update Custom Metrics
- [ ] List Goals
- [ ] Insert / Update Goals
- [ ] Differentiate between GA Standard and GA Premium and adapt datavalidation to it

### Bugs
- [x] Verify positions of dataRange.clearContent()
- [x] Refactor sheet init functions
- [x] Refactor API init functions
- [ ] Make sure onChange validation keeps working when a column gets deleted
- [ ] Verify onChange validation to validate all rows with changes
- [ ] Before running datavalidation, verify the columns / sheet structure


## Installation

> When this library was developed, write operations to the Google Analytics Management API where still in a limited beta available as developer preview. See [this page](https://developers.google.com/analytics/devguides/config/mgmt/v3/account-management) for more information and how to request access to the beta.

## Credits
This project is based on this repository: https://github.com/nicholasblexrud/ga-auditor/