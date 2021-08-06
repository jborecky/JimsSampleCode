#!/bin/bash
#
#===============================================================================================================
#        NAME: EOD_snap_shutdown_vms.sh
#      AUTHOR: Jim Borecky - Homelab
#        DATE: 06/12/2019
# DESCRIPTION: Tanium sensor to check the health of the Crowdstrike agent on MacOS.
#---------------------------------------------------------------------------------------------------------------
# Version   Date            Initial             Comment
#---------------------------------------------------------------------------------------------------------------
# 1.0       06/012/2019     Jim Borecky         Original
#
# 
#
#===============================================================================================================

#Format timestamp to append to File name
timestamp=$(date '+%D_%H:%M')
timestamp=${timestamp//\//-}

#Grab any machines that is running and take snapshot then send shutdown command
for i in $(virsh list | grep running | awk '{print $2}'); do
	echo Taking snapshot of $i
	virsh snapshot-create-as --domain $i --name "EOD_$timestamp"
	virsh shutdown $i
done

#Continue to loop until all the machines are shutdown
count=1
while (( $count > 0 ))
do
  count=0
  #Give them some time to shutdown, but not too much.
  echo Sleeping for 1 minutes $(date)
  sleep 1m

  #Let figure out how many still are runnning and send the shutdown command again.
  for vm in $(virsh list | grep running | awk '{print $2}'); do
     	echo Shutting down $vm
	virsh shutdown $vm
        count=$((count+1))
	echo Machines still running=$count
	stillrunning=$true
  done
done
