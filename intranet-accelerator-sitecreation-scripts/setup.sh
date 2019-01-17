#!/usr/bin/env bash

set -o errexit
set -o pipefail
set -o nounset

# helper functions
. ./_functions.sh

# Prerequisites
msg 'Checking prerequisites...'

set +e
_=$(command -v o365);
if [ "$?" != "0" ]; then
  error 'ERROR'
  echo
  echo "You don't seem to have the Office 365 CLI installed."
  echo "Install it by executing 'npm i -g @pnp/office365-cli'"
  echo "More information: https://aka.ms/o365cli"
  exit 127
fi;

_=$(command -v jq);
if [ "$?" != "0" ]; then
  error 'ERROR'
  echo
  echo "You don't seem to have jq installed."
  echo "Install it from https://stedolan.github.io/jq/"
  exit 127
fi;
set -e
success 'DONE'

# default args values
tenantUrl="https://spfxdeveloper1.sharepoint.com"
# For classic site mention owners mandatory
siteAdmin="rabia@spfxdeveloper1.onmicrosoft.com"

# script arguments
while [ $# -gt 0 ]; do
  case $1 in
    -t|--tenantUrl)
      shift
      tenantUrl=$1
      ;; 
    -m|--mainTitle)
      shift
      mainTitle=$1
      ;;
    -h|--help)
      help
      exit
      ;;
    *)
      error "Invalid argument $1"
      exit 1
  esac
  shift
done

if [ -z "$tenantUrl" ]; then
  error 'Please specify tenant URL e.g https://tenant.sharepoint.com'
  echo
  help
  exit 1
fi

createCommunicationSite() {
 site=$(o365 spo site get --url "$1" --output json || true)
  if $(isError "$site"); then
    msg "Creating communication site at $1 "
    sub '- Creating site...'
    o365 spo site add --type CommunicationSite --url "$1" --title "$3" --alias "$2" 
    sub '- Applying custom theme at Communication site'
    o365 spo theme apply --name "Mirvac Blue"  --webUrl "$1"
    success 'DONE'
  else
    warning 'EXISTS'
  fi
}
#Set custom theme
  msg 'Applying custom themes'
  o365 spo theme set --name "Mirvac Blue" --filePath './mirvactheme.json'
  success 'DONE'
 
# Provision main site
intranetUrl=$tenantUrl/sites/intranet-05
msg "Creating main site at $intranetUrl"
  site=$(o365 spo site get --url $intranetUrl --output json || true)
  if $(isError "$site"); then
  sub '- Creating main intranet site'
    o365 spo site add --type CommunicationSite --url $intranetUrl --title "Intranet Accelerator" --description 'Intranet Accelerator'  
  sub '- Applying custom theme at main intranet site'
    o365 spo theme apply --name "Mirvac Blue"  --webUrl $intranetUrl
    success 'DONE'
  else
    warning 'EXISTS'
  fi 
echo
success "Intranet portal has been successfully provisioned to $intranetUrl"

#Register main site as hubsite
sub "- Registering $intranetUrl hub site..."
out=$(o365 spo hubsite register --url $intranetUrl --output json || true)     
success 'DONE'

# Provision Communication sites
intranetContentUrl=$tenantUrl/sites/intranet-content-05
intranetNewsUrl=$tenantUrl/sites/intranet-news-05
intranetProjectsUrl=$tenantUrl/sites/intranet-projects-05
intranetServicesUrl=$tenantUrl/sites/intranet-services-05
intranetTemplatesUrl=$tenantUrl/sites/intranet-templates-05


msg "Provisioning communication  site at $intranetContentUrl..."
createCommunicationSite $intranetContentUrl "intranet-content-05" "Content Hub"
msg "Provisioning communication  site at $intranetNewsUrl..."
createCommunicationSite $intranetNewsUrl "intranet-news-05" "News Hub"
msg "Provisioning communication  site at $intranetProjectsUrl..."
createCommunicationSite $intranetProjectsUrl "intranet-projects-05" "Projects Hub"
msg "Provisioning communication  site at $intranetServicesUrl..."
createCommunicationSite $intranetServicesUrl "intranet-services-05" "Service Hub"
msg "Provisioning communication  site at $intranetTemplatesUrl..."
createCommunicationSite $intranetTemplatesUrl "intranet-templates-05" "Templates"

# Search classic site
  intranetSearchUrl=$tenantUrl/sites/intranet-search-05
  
  site=$(o365 spo site get --url $intranetSearchUrl --output json || true)
  msg "Provisioning classic search site at $intranetSearchUrl..."
   if $(isError "$site"); then
   sub 'Creating search site..'
    o365 spo site classic add --url $intranetSearchUrl --owner $siteAdmin --title "Search"  --webTemplate "SRCHCEN#0" --timeZone 76 --wait 
    success 'DONE'
  else
    warning 'EXISTS'
  fi

# Check if app catalog is already provisioned in the tenancy, if not provision it
msg 'Retrieving tenant app catalog URL...'
appCatalogUrl=$(o365 spo tenant appcatalogurl get)
if [ -z "$appCatalogUrl" ]; then
  appCatalogUrl=$tenantUrl/sites/apps
  error "Couldn't retrieve tenant app catalog"
  msg "Provisioning an app catalog at $tenantUrl/sites/apps"
  o365 spo site appcatalog add --url $appCatalogUrl
  exit 1
  else
  msg "App catalog found at $appCatalogUrl"
fi
success 'DONE'

AssociateHubSite()
{
  sub "Associate $1 hub to hubsite"
  o365 spo hubsite connect --url $1 --hubSiteId $2
  success 'DONE'
}

#Associate sites to hubsite
  msg 'Associate all communication sites to hubsite'
  sub '- Retrieving Hubsite ID to associate sites to hubsite...'
  siteId=$(o365 spo site get --url $intranetUrl --output json | jq -r '.Id')

  msg 'Associating sites to Hubsite...' 
  AssociateHubSite  $intranetContentUrl  "$siteId"
  AssociateHubSite  $intranetNewsUrl  "$siteId"
  AssociateHubSite  $intranetProjectsUrl  "$siteId"
  AssociateHubSite  $intranetTemplatesUrl  "$siteId"
  AssociateHubSite  $intranetServicesUrl  "$siteId"


