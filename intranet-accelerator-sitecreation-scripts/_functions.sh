help() {
  echo
  echo "Intranet Accelerator  sites setup script"
  echo
  echo "Usage: ./setup.sh [options]"
  echo
  echo "Options:"
  echo
  echo "--help                           Output usage information"
  echo "-t, --tenantUrl <tenantUrl>      URL of the tenant to provision E.g https://tenant.sharepoint.com"  
  echo "-m, --mainTitle [company]          Name of the client to use in the provisioned sites. Default 'E2'"
  echo
}

isError() {
  # some error messages can have line breaks which break jq, so they need to be
  # removed before passing the string to jq
  res=$(echo "${1//\\r\\n/ }" | jq -r '.message')
  if [[ -z "$res" || "$res" = "null" ]]; then return 1; else return 0; fi
}

msg() {
  printf -- "$1"
}

sub() {
  printf -- "\033[90m$1\033[0m"
}

warningMsg() {
  printf -- "\033[33m$1\033[0m"
}

success() {
  printf -- "\033[32m$1\033[0m\n"
}

warning() {
  printf -- "\033[33m$1\033[0m\n"
}

error() {
  printf -- "\033[31m$1\033[0m\n"
}

errorMessage() {
  # some error messages can have line breaks which break jq, so they need to be
  # removed before passing the string to jq
  msg=$(echo "${1//\\r\\n/ }" | jq -r ".message")
  error "$msg"
}

# $1 string with key-value pairs
# $2 name of the property for which to retrieve value
getPropertyValue() {
  echo "$1" | grep -o "$2:\"[^\"]\\+" | cut -d"\"" -f2
}
