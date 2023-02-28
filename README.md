## project-client-helpers

This is where you include your WebPart documentation.

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

### Compatibility
- node: 8.17.0
- npm: 6.13.4
- gulp-cli: 2.3.0
- yo: 2.0.6
- @microsoft/generator-sharepoint: 1.10.0
- Fabric UI: 5.132.0
- React: 15.6.2
- React-DOM: 15.6.2
- office-ui-fabric-react: 5.135.6
- @microsoft/sp-http: 1.12.1
