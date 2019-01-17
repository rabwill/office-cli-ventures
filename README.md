# office-cli-ventures

## Prerequisites

- Install
  - [Office 365 CLI](https://aka.ms/o365cli) latest beta (`npm i -g @pnp/office365-cli@next`)
  - [jq](https://stedolan.github.io/jq/) (Use [Chocolatey NuGet](https://chocolatey.org/) to install jq 1.5 with
      ```
      chocolatey install jq
      ```
- Configure
  - Set the user who will run the setup script in setup.sh line 38 as siteAdmin
  - Set the urls of all the sites provisioned as communication sites
  ```
  intranetContentUrl=$tenantUrl/sites/intranet-content-05
  intranetNewsUrl=$tenantUrl/sites/intranet-news-05
  intranetProjectsUrl=$tenantUrl/sites/intranet-projects-05
  intranetServicesUrl=$tenantUrl/sites/intranet-services-05
  intranetTemplatesUrl=$tenantUrl/sites/intranet-templates-05
  ```
- Execute
  - `o365 spo login [tenant admin]`, eg. `o365 spo connect https://tenant-admin.sharepoint.com` in a new bash than the one to execute the provisioning. Keep this open.


## Setup
- Go to the folder location where `setup.sh`is in terminal
-  run `chmod +x ./setup.sh`
- Execute the provisioning script by running below command in the cmd line
```
./setup.sh --tenantUrl https://tenant.sharepoint.com
```

Following are the options you can pass to the script:

argument|description|required|default value|example value
--------|-----------|--------|-------------|-------------
`-t, --tenantUrl`|URL of the SharePoint tenant where the sites have to be provisioned
|yes|`https://tenant.sharepoint.com`|`https://tenant.sharepoint.com`
`
