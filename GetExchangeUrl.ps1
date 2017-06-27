###############################################################################################################################################################################
###                                                                                                           																###
###		.INFORMATIONS																																						###
###  	Script by Drago Petrovic -                                                                            																###
###     Technical Blog -               https://msb365.abstergo.ch                                               															###
###     GitHub Repository -            https://github.com/MSB365                                          	  																###
###     Webpage -                                                                  																							###
###     Xing:				   		   https://www.xing.com/profile/Drago_Petrovic																							###
###     LinkedIn:					   https://www.linkedin.com/in/drago-petrovic-86075730																					###
###																																											###
###		.VERSION																																							###
###     Version 1.0 - 13/04/2017                                                                              																###
###     Version 2.0 -                                                                               																		###
###     Revision -                                                                                            																###
###                                                                                                           																### 
###               v1.0 - Initial script										                                  																###
###               				                                          																									###
###																																											###
###																																											###
###		.SYNOPSIS																																							###
###		GetExchangeUrl.ps1																																					###
###																																											###
###		.DESCRIPTION																																						###
###		Script to allow you to get all virtual directories URLs							.																					###
###																																											###
###		.PARAMETER																																							###
###																																											###
###																																											###
###		.EXAMPLE																																							###
###		.\GetExchangeUrl.ps1																																				###
###																																											###
###		.NOTES																																								###
###																																											###
###																																											### 	
###																																											###
###																																											###
###                                                                                                           																###  	
###     .COPIRIGHT                                                            																								###
###		Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 					###
###		to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 					###
###		and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:							###
###																																											###
###		The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.										###
###																																											###
###		THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 				###
###		FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 		###
###		WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.			###
###                 																																						###
###                                                																															###
###                                                                                                           																###
###                                                                                                           																###
###############################################################################################################################################################################
#
# Variables 

$localserver = $env:computerName
$ExOrgCfg = Get-OrganizationConfig


write-host “” 
Write-host “This script will get all Vdir URLS for the local server or all Exchange servers in the organization.” -foregroundcolor Yellow 
Write-host “Keep it simple but significant” -foregroundcolor magenta
Write-host “” 

    
###################################### 
# Validate if a cert information should also be pulled from servers.
#[string]$set = Read-host “Do you want to pull the certificates installed? (Y/N)” 
Write-host “” 

if ($set -eq “Y”)    { 
    [boolean]$certInfo = $true
}    else    { 
    [boolean]$certInfo = $false
} 

###################################### 
# Validate if an operation is delivered to all CAS servers or only to local server 
# 
[string]$isglobal = Read-host “Do you want to get vdirs urls for all CAS servers (Y) or for local server only (N)? (Y/N)” 
Write-host “” 

if ($isglobal -eq “Y”)
{
    [array]$exchServers = Get-ExchangeServer | ?{$_.ServerRole -Match "ClientAccess"}
} else {
    [array]$exchServers = Get-ExchangeServer -identity $localserver
}

###################################### 
# Validate if results should be output to CSV
[string]$set = Read-host “Do you want output the results? (Y/N)” 
Write-host “” 

if ($set -like “Y”) { 
    [boolean]$output = $true
}    else    { 
    [boolean]$output = $false
}

if ($output -eq $true) {
    [string]$outfile = $env:USERPROFILE + "\Desktop\CAS_VDIR_URLs.csv"
    Write-host “” 

    $outArray = @()
    
}


Foreach ($server in $exchServers) {

$temp = New-Object System.Object
$temp | Add-Member -type NoteProperty -name Server -value $server.Name

######################################
# Get Autodiscover SCPs
#
#Foreach ($server in $exchServers) {

    Write-host “Getting Autodiscover Service Connection Point” -foregroundcolor Yellow 
    write-host “” 

    $SCPCurrent = Get-ClientAccessServer -identity $server.Name

    Write-host “Looking at Server: ” $server.Name
    Write-host “Current SCP value: ” $SCPCurrent.AutoDiscoverServiceInternalUri.absoluteuri
    write-host “”
    write-host “”

    if ($outFile) {
        
        $temp | Add-Member -type NoteProperty -name AutodiscoverSCP -value $SCPCurrent.AutoDiscoverServiceInternalUri.absoluteuri

    }

#} 

######################################
# Get OAB URLs
#
    Write-host “Getting OAB Virtual Directories” -foregroundcolor Yellow 
    write-host “” 
    
    $OABCurrent = Get-OABVirtualDirectory -server $server.Name -ADPropertiesOnly
      
    Write-host “Looking at Server: ” $server.Name
    Write-host “Current Internal Value: ” $OABCurrent.internalURL
    Write-host “Current External Value: ” $OABCurrent.externalURL
    write-host “”
    write-host “”

    if ($outFile) {
        
        $temp | Add-Member -type NoteProperty -name OABInternal -value $OABCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name OABExternal -value $OABCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name OABRequireSSL -value $OABCurrent.requireSSL

    }

#}


######################################
# Get EWS URLs
#
#Foreach ($server in $exchServers) {
    
    Write-host “Getting Exchange Web Services Virtual Directories” -foregroundcolor Yellow 
    write-host “” 

    [array]$EWSCurrent = Get-WebServicesVirtualDirectory -server $server.Name -ADPropertiesOnly

    Write-host “Looking at Server: ” $server.Name
    Write-host “Current Internal Value: ” $EWSCurrent.internalURL 
    Write-host “Current External Value: ” $EWSCurrent.externalURL 
    write-host “”
    write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name EWSInternal -value $EWSCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name EWSExternal -value $EWSCurrent.externalURL
    }

#}


######################################
# Get EMS URLs
#
#Foreach ($server in $exchServers) {

    Write-host “Getting Exchange Management Shell Virtual Directories” -foregroundcolor Yellow 
    write-host “” 

    $EMSCurrent = Get-PowerShellVirtualDirectory -server $server.Name -ADPropertiesOnly

    Write-host “Looking at Server: ” $server.Name
    Write-host “Current Internal Value: ” $EMSCurrent.internalURL 
    Write-host “Current External Value: ” $EMSCurrent.externalURL 
    write-host “”
    write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name EMSInternal -value $EMSCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name EMSExternal -value $EMSCurrent.externalURL
    }

#}


######################################
# Get ECP URLs
#
#Foreach ($server in $exchServers) {

    Write-host “Getting ECP Virtual Directories” -foregroundcolor Yellow 
    write-host “” 

    $ECPCurrent = Get-ECPVirtualDirectory -server $server.name -ADPropertiesOnly

    Write-host “Looking at Server: ” $server.name
    Write-host “Current Internal Value: ” $ECPCurrent.internalURL 
    Write-host “Current External Value: ” $ECPCurrent.externalURL 
    write-host “”
    write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name ECPInternal -value $ECPCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name ECPExternal -value $ECPCurrent.externalURL
    }

#}


###################################### 
# Get OWA URLs 
#
#Foreach ($server in $exchServers) {
    Write-host “Getting OWA Virtual Directories” -foregroundcolor Yellow 
    write-host “” 

    $OWACurrent = Get-OWAVirtualDirectory -server $server.Name -ADPropertiesOnly | ? {$_.name -like "*owa*"}

    Write-host “Looking at Server: ” $server.Name
    Write-host “Current Internal Value: ” $OWACurrent.internalURL 
    Write-host “Current External Value: ” $OWACurrent.externalURL 
    write-host “”
    write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name OWAInternal -value $OWACurrent.internalURL
        $temp | Add-Member -type NoteProperty -name OWAExternal -value $OWACurrent.externalURL
    }

#} 


###################################### 
# Get EAS URLs
#
#Foreach ($server in $exchServers) {
    
    Write-host “Getting EAS Virtual Directories” -foregroundcolor Yellow 
    write-host “” 

    $EASCurrent = Get-ActiveSyncVirtualDirectory -server $server.name -ADPropertiesOnly

    Write-host “Looking at Server: ” $server.name
    Write-host “Current Internal Value: ” $EASCurrent.internalURL 
    Write-host “Current External Value: ” $EASCurrent.externalURL 
    write-host “”
    write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name EASInternal -value $EASCurrent.internalURL
        $temp | Add-Member -type NoteProperty -name EASExternal -value $EASCurrent.externalURL
    }

#} 

###################################### 
# Get OutlookAnywhere URLs
#	
#Foreach ($server in $exchServers) {
		
        Write-host “Getting Outlook Anywhere hostnames” -foregroundcolor Yellow 
		write-host “” 

        $OACurrent = Get-OutlookAnywhere -server $server.Name -ADPropertiesOnly | Select @{l='IISAuthenticationMethods';e={[string]::join(" ", ($_.IISAuthenticationMethods))}},`
        InternalHostname, InternalClientsRequireSsl, InternalClientAuthenticationMethod, externalhostname, ExternalClientsRequireSsl, ExternalClientAuthenticationMethod
		
        Write-host “Looking at Server: ” $server.Name
        Write-host “Current IIS Auth Methods: ” $OACurrent.IISAuthenticationMethods
    	Write-host “Current Internal Value: ” $OACurrent.InternalHostname 
        Write-host “Current Internal Clients RequireSSL: ” $OACurrent.InternalClientsRequireSsl
        Write-host “Current Internal Auth Methods: ” $OACurrent.InternalClientAuthenticationMethod
    	Write-host “Current External Value: ” $OACurrent.externalhostname 
        Write-host “Current External Clients RequireSSL: ” $OACurrent.ExternalClientsRequireSsl
        Write-host “Current External Auth Methods: ” $OACurrent.ExternalClientAuthenticationMethod
        write-host “”
        write-host “”

    if ($outFile) {

        $temp | Add-Member -type NoteProperty -name OAIISMethods -value $OACurrent.IISAuthenticationMethods
        $temp | Add-Member -type NoteProperty -name OAInternal -value $OACurrent.InternalHostname
        $temp | Add-Member -type NoteProperty -name OAInternalSSL -value $OACurrent.InternalClientsRequireSsl
        $temp | Add-Member -type NoteProperty -name OAInternalAuthMethod -value $OACurrent.InternalClientAuthenticationMethod
        $temp | Add-Member -type NoteProperty -name OAExternal -value $OACurrent.externalhostname
        $temp | Add-Member -type NoteProperty -name OAExternalSSL -value $OACurrent.ExternalClientsRequireSsl
        $temp | Add-Member -type NoteProperty -name OAExternalAuthMethod -value $OACurrent.ExternalClientAuthenticationMethod
    }

#} 


###################################### 
# Get MAPI URLs (Only for Exchange 2013)
# MAPI Requires: 
# - .Net 4.5.2 to be installed on CAS
# -  Requires COMPLUS_DisableRetStructPinning enviroment variable to be set
# -  Requres MAPI to be enabled on organizational config

if ($server.IsE15OrLater -eq $true) {
if ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15) {

	if ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Build -ge 847){

        if ($ExOrgCfg.MapiHTTPEnabled -eq $false) {

            Write-host “Your organization supports MAPI over HTTP, however it is currently not enabled.” -foregroundcolor Yellow
            write-host “”
            [string]$set = Read-host “Do you wish to view the current MAPI VDIR URLs (Y/N)”
            write-host “”
        }

        if ($set -eq "Y" -or $ExOrgCfg.MapiHTTPEnabled -eq $true) {
        
            #Foreach ($server in $exchServers) {
                


			        Write-host “Getting MAPI Virtual Directories” -foregroundcolor Yellow 
			        write-host “” 

                    $MAPICurrent = Get-MAPIVirtualDirectory -server $server.Name -ADPropertiesOnly | Select @{l='IISAuthenticationMethods';e={[string]::join(" ", ($_.IISAuthenticationMethods))}},`
                    internalURL, externalURL, @{l='InternalAuthenticationMethods';e={[string]::join(" ", ($_.InternalAuthenticationMethods))}},`
                    @{l='ExternalAuthenticationMethods';e={[string]::join(" ", ($_.ExternalAuthenticationMethods))}}

			        Write-host “Looking at Server: ” $server.Name
                    Write-host “Current IIS Auth Methods: ” $MAPICurrent.IISAuthenticationMethods
    		        Write-host “Current Internal Value: ” $MAPICurrent.internalURL
    		        Write-host “Current Internal Auth Methods: ” $MAPICurrent.InternalAuthenticationMethods
    		        Write-host “Current External Value: ” $MAPICurrent.externalURL
                    Write-host “Current External Auth Methods: ” $MAPICurrent.ExternalAuthenticationMethods
                    write-host “”
                    write-host “”

                    if ($outFile) {

                        $temp | Add-Member -type NoteProperty -name MAPIIISAuthMethods -value $MAPICurrent.IISAuthenticationMethods
                        $temp | Add-Member -type NoteProperty -name MAPIInternal -value $MAPICurrent.InternalURL
                        $temp | Add-Member -type NoteProperty -name MAPIInternalAuthMethods -value $MAPICurrent.InternalAuthenticationMethods
                        $temp | Add-Member -type NoteProperty -name MAPIExternal -value $MAPICurrent.ExternalURL
                        $temp | Add-Member -type NoteProperty -name MAPIExternalAuthMethods -value $MAPICurrent.ExternalAuthenticationMethods
                    }

                #}

            }
    			   
        }
	
    }

    
}

$outArray += $temp

}

if ($outFile) {
    #$output += $temp
    $outArray | Export-Csv $outfile -NoTypeInformation -Force
    Write-Host "Output file is $outfile"
}