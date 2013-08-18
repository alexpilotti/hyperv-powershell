#REQUIRES -Version 2.0

# Copyright 2013 Cloudbase Solutions Srl
# All Rights Reserved.
#
#    Licensed under the Apache License, Version 2.0 (the "License"); you may
#    not use this file except in compliance with the License. You may obtain
#    a copy of the License at
#
#         http://www.apache.org/licenses/LICENSE-2.0
#
#    Unless required by applicable law or agreed to in writing, software
#    distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
#    WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
#    License for the specific language governing permissions and limitations
#    under the License.

#Based on: http://blogs.technet.com/b/m2/archive/2010/07/29/how-to-get-the-ip-address-of-a-virtual-machine-from-hyper-v.aspx

#Usage: Import-Module VMKVP.psm1

filter Import-CimXml {
    $cimXml = [Xml]$_
    $cimObj = New-Object -TypeName System.Object
    foreach ($cimProperty in $cimXml.SelectNodes("/INSTANCE/PROPERTY")) {
        if ($cimProperty.Name -eq "Name" -or $cimProperty.Name -eq "Data") {
            $cimObj | Add-Member -MemberType NoteProperty -Name $cimProperty.NAME -Value $cimProperty.VALUE
        }
    }
    $cimObj
}

function GetVMKVP($host, $filter) {
    if([Environment]::OSVersion.Version -ge (new-object 'Version' 6, 2)) {
        $ns = "root\virtualization\v2"
    }
    else {
        $ns = "root\virtualization"
    }

    $vm = Get-WmiObject -Class Msvm_ComputerSystem -Namespace $ns -ComputerName $host -Filter ("Description <> 'Microsoft Hosting Computer System' AND " + $filter)
    if(!$vm) {
        throw "Virtual machine not found"
    }
    if($vm.EnabledState -ne 2) {
        throw "The virtual machine """ + $vm.ElementName + """ is not running"
    }

    $kvp = Get-WmiObject -Namespace $ns -Query "Associators of {$vm} Where AssocClass=Msvm_SystemDevice ResultClass=Msvm_KvpExchangeComponent"
    $kvp.GuestIntrinsicExchangeItems | Import-CimXml
}

<#
.DESCRIPTION
    This Cmdlet retrieves the key value pairs (KVP) for a given VM
.NOTES
    Copyright 2013 - Cloudbase Solutions Srl
.LINK
    http://www.cloudbase.it
.EXAMPLE
    Get-VMKVP MyVM
.EXAMPLE
    Get-VM | where {$_.State -eq "Running"} | Get-VMKVP
#>
function Get-VMKVP {
    [CmdletBinding(DefaultParameterSetName="VMName")]
    param (

        [parameter(Mandatory=$true,Position=0,ParameterSetName="VMName")]
        [string[]]$VMName,

        [parameter(Mandatory=$false,Position=0,ParameterSetName="VM", ValueFromPipeline=$true)]
        [PSObject[]]$VM,

        [parameter(Mandatory=$false,Position=1)]
        [string]$HyperVHost = "127.0.0.1"
    )
    PROCESS {
        if($VM) {
            $name = $VM.Id
            GetVMKVP $HyperVHost "Name = '$name'" $false $Wait
        }
        else {
            foreach($ElementName in $VMName) {
                GetVMKVP $HyperVHost "ElementName = '$ElementName'" $false $Wait
            }
        }
    }
}
