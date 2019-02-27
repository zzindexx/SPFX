## Phonebook with ogranizational structure
This webpart allows you to create a simple phonebook for your organization with only one webpart.
![overview](/assets/overview.png)

### Prerequisites
This webpart expects that managed metadata service and search service applications are created in the farm.
Also some additional configuration required:
1. Special termset, that consists of reused terms from Department termset and represents organizational structure
2. Configure managed properties for search
    1. Make "PrefferedName" managed property searchable
        ```powershell
        $ssa = Get-SPEnterpriseSearchServiceApplication
        $mp = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $ssa -Identity "PreferredName"
        $mp.Sortable = $true
        $mp.Update()
        ```
    2. Create a new managed property "DepartmentTaxId", make it querable and map it to "ows_taxId_SPShDepartment" crawled property
3. Change the url of your root SharePoint site in externals section of config/config.json file

### Components used
* [rc-tree](https://www.npmjs.com/package/rc-tree)
* [react-js-pagination](https://www.npmjs.com/package/react-js-pagination)


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

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO