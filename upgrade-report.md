# Upgrade project spfx-react-directory to v1.17.4

Date: 11/17/2023

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.17.4. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.17.4
```

File: [./package.json:17:5](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.17.4
```

File: [./package.json:18:5](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.17.4
```

File: [./package.json:19:5](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.17.4
```

File: [./package.json:21:5](./package.json)

### FN001021 @microsoft/sp-property-pane | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane

Execute the following command:

```sh
npm i -SE @microsoft/sp-property-pane@1.17.4
```

File: [./package.json:20:5](./package.json)

### FN001034 @microsoft/sp-adaptive-card-extension-base | Optional

Install SharePoint Framework dependency package @microsoft/sp-adaptive-card-extension-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-adaptive-card-extension-base@1.17.4
```

File: [./package.json:14:3](./package.json)

### FN002022 @microsoft/eslint-plugin-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-plugin-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-plugin-spfx@1.17.4
```

File: [./package.json:34:5](./package.json)

### FN002023 @microsoft/eslint-config-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-config-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-config-spfx@1.17.4
```

File: [./package.json:33:5](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.17.4
```

File: [./package.json:36:5](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.17.4
```

File: [./package.json:37:5](./package.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.17.4"
  }
}
```

File: [./.yo-rc.json:10:5](./.yo-rc.json)

### FN001035 @fluentui/react | Required

Install SharePoint Framework dependency package @fluentui/react

Execute the following command:

```sh
npm i -SE @fluentui/react@7.199.1
```

File: [./package.json:14:3](./package.json)

### FN002020 @microsoft/rush-stack-compiler-4.5 | Required

Upgrade SharePoint Framework dev dependency package @microsoft/rush-stack-compiler-4.5

Execute the following command:

```sh
npm i -DE @microsoft/rush-stack-compiler-4.5@0.5.0
```

File: [./package.json:32:22](./package.json)

### FN014008 Hosted workbench type in .vscode/launch.json | Recommended

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration

```json
{
  "configurations": [
    {
      "type": "msedge"
    }
  ]
}
```

File: [.vscode\launch.json:8:7](.vscode\launch.json)

### FN002024 eslint | Required

Install SharePoint Framework dev dependency package eslint

Execute the following command:

```sh
npm i -DE eslint@8.7.0
```

File: [./package.json:32:3](./package.json)

### FN007002 serve.json initialPage | Required

Update serve.json initialPage URL

```json
{
  "initialPage": "https://{tenantDomain}/_layouts/workbench.aspx"
}
```

File: [./config/serve.json:5:3](./config/serve.json)

### FN010010 .yo-rc.json @microsoft/teams-js SDK version | Recommended

Update @microsoft/teams-js SDK version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "2.9.1"
    }
  }
}
```

File: [./.yo-rc.json:2:3](./.yo-rc.json)

### FN014009 Hosted workbench URL in .vscode/launch.json | Recommended

In the .vscode/launch.json file, update the url property for the hosted workbench launch configuration

```json
{
  "configurations": [
    {
      "url": "https://{tenantDomain}/_layouts/workbench.aspx"
    }
  ]
}
```

File: [.vscode\launch.json:8:7](.vscode\launch.json)

### FN015009 config\sass.json | Required

Add file config\sass.json

Execute the following command:

```sh
cat > "config\sass.json" << EOF 
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/sass.schema.json"
}
EOF
```

File: [config\sass.json](config\sass.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @microsoft/sp-core-library@1.17.4 @microsoft/sp-lodash-subset@1.17.4 @microsoft/sp-office-ui-fabric-core@1.17.4 @microsoft/sp-webpart-base@1.17.4 @microsoft/sp-property-pane@1.17.4 @microsoft/sp-adaptive-card-extension-base@1.17.4 @fluentui/react@7.199.1
npm i -DE @microsoft/eslint-plugin-spfx@1.17.4 @microsoft/eslint-config-spfx@1.17.4 @microsoft/sp-build-web@1.17.4 @microsoft/sp-module-interfaces@1.17.4 @microsoft/rush-stack-compiler-4.5@0.5.0 eslint@8.7.0
npm dedupe
cat > "config\sass.json" << EOF 
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/sass.schema.json"
}
EOF
```

### Modify files

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.17.4"
  }
}
```

Update @microsoft/teams-js SDK version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "2.9.1"
    }
  }
}
```

#### [.vscode\launch.json](.vscode\launch.json)

In the .vscode/launch.json file, update the type property for the hosted workbench launch configuration:

```json
{
  "configurations": [
    {
      "type": "msedge"
    }
  ]
}
```

In the .vscode/launch.json file, update the url property for the hosted workbench launch configuration:

```json
{
  "configurations": [
    {
      "url": "https://{tenantDomain}/_layouts/workbench.aspx"
    }
  ]
}
```

#### [./config/serve.json](./config/serve.json)

Update serve.json initialPage URL:

```json
{
  "initialPage": "https://{tenantDomain}/_layouts/workbench.aspx"
}
```
