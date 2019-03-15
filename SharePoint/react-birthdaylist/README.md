## Birthday List
This webpart allows you to show a list of birthdays
![overview](https://github.com/zzindexx/SPFX/blob/master/SharePoint/react-birthdaylist/assets/overview.png)

### Prerequisites
This webpart expects that managed metadata service and search service applications are created in the farm.
Also some additional configuration required:
    1. Create a new managed property for stoding birth date (default is Birthday), make it querable and sortable, and map it to "People:SPS-Birthday" crawled property


### Properties

#### Additional query
You can specify additional query to filter birthday list. For example, you can configure web part to show birthdays only for one department
![additionalQuery](https://github.com/zzindexx/SPFX/blob/master/SharePoint/react-birthdaylist/assets/additionalquery.gif)

#### Display type
The web part can show list for today, current week and current month.
![displaytypes](https://github.com/zzindexx/SPFX/blob/master/SharePoint/react-birthdaylist/assets/displaytypes.gif)


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
