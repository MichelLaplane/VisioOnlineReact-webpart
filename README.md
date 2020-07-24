## Visio-online-react-webpart
This project implement a SPFx WebPart using VisioOnline JS API.
It gives you some guideline to play with Visio Online (Visio for the web) inside a SharePoint Page.

You can use it directly for creating a sharePoint Package that you can installa on your M365 tenant.

You can also run it in Visual studio Code in debug mose using the Online Workbench of your M365 tenant

Screen shot 1
![alt text](https://user-images.githubusercontent.com/15141659/88371866-243cb180-cd95-11ea-8c1c-c46b24c3d7b8.png)


Screen shot 2
![alt text](https://user-images.githubusercontent.com/15141659/88372590-82b65f80-cd96-11ea-9875-1d5f6929ae88.png)


Screen shot 3
![alt text](https://user-images.githubusercontent.com/15141659/88372609-89dd6d80-cd96-11ea-8ffc-b83197dc90cd.png)

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
