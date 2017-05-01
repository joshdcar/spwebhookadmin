
## webhook-manager-webpart

A SharePoint Webhook administration webpart built using the SharePoint Framework (SPFX) and React. 

NOTE: This webpart is a work in progress.  Please report any issues and I will work to get them resolved when time permits. Contributions are also welcomed.  More documentation forthcoming..

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

###Build options

gulp clean 
gulp test 
gulp serve 
gulp bundle
gulp package-solution 

** See SharePoint Framwork for additional guidance



