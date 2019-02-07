## office-ui-control in SPFx webpart  

## Summary
The SPFx sample solution to demostrate CRUD operation functionality along with utilization of Office Febric UI Components Control into SPFx webpart.

This Solution is contains following [Office-UI-Febric Components](https://developer.microsoft.com/en-us/fabric#/components) and @pnp/sp framework:

* Basic Input Component(Textbox, button)
* People Picker Control
* Dropdown Control
* Details List Component with Sorting and Filtering Functionality.

## Applies to

* [SharePoint Framework](https:/dev.office.com/sharepoint)
* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)


### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

* gulp clean - TODO
* gulp test - TODO
* gulp serve - TODO
* gulp bundle - TODO
* gulp package-solution - TODO
