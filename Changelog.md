# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [0.1.3] - 2023-08-03

### Fixed

- Fixed the way that the Chrome driver is loaded in order to support Chrome 115.0  
This will require that the selenium dependency is updated to version 4.11.2+
- Fixed some errors due to changes in the Gainsight timeline html elements

## [0.1.2] - 2022-03-18

### Fixed

- Fixed an issue related with the UI element changes on the Azure login page
- Fixed an issue related with the incorrect parsing of dates from the excel file

### Changed

- Improved the logic on click elements to wait for them to be clicable before trying to click
- Refactored the selectors to dedicated variables in order to reduce redundancy

## [0.1.1] - 2022-02-28

### Fixed

- Fixed an issue related with leading zeros when trying to insert dates

## [0.1] - 2022-02-02

### Added

- Initial release
