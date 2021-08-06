#!/bin/bash
#
#===============================================================================================================
#        NAME: CrowdstrikeHealth.sh
#      AUTHOR: Jim Borecky - Homelab
#        DATE: 08/06/2021
# DESCRIPTION: Tanium sensor to check the health of the Crowdstrike agent on MacOS.
#---------------------------------------------------------------------------------------------------------------
# Version   Date            Initial             Comment
#---------------------------------------------------------------------------------------------------------------
# 1.0       08/06/2021      Jim Borecky         Original
#
# 
#
#===============================================================================================================

sep'|'
# Current windows columns to match
#"Installed"    "Install Display Name"  "Version"   "Date Installed"    "Service Disply Name"   "Status"    "Service Type"

#check to see if installed
if [ -f "/Application/Falcon.app/Contents/Resources/falconctl"]; then
    #verifying the agent is connected to crowdstrike
    Status=$(/Application/Falcon.app/Contents/Resources/falconctl stats | grep -i State: | sed -E 's/^.*: //' | sed -e 's/[[:space:]]*//')
    Version=$(/Application/Falcon.app/Contents/Resources/falconctl stats | grep -i version: | sed -E 's/^.*: //' | sed -e 's/[[:space:]]*//')
    ServiceName=$(sudo launchctl list | grep -i crowd | awk '{print $3}' | sed -e 's/[[:space:]]*//')
    DateInstalled=$(sudo softwareupdate --history --all | grep -i Falcon | sed 's/^.*       //' | sed 's [Ii]nstaller//' | sed -e 's/[[:space:]]*//')
    DisplayName=$(sudo softwareupdate --history --all | grep i Falcon | sed 's/ .*//' | sed -e 's/[[:space:]]*//')

    echo "True" $sep $DisplayName $sep $Version $sep $DateInstalled $sep $sep $ServiceName $sep $Status $sep "N/A"
else
    echo "False||||Crowdstrike service not found|||"
fi