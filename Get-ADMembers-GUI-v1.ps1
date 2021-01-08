
# ========================================================
#
# 	Script Information
#
#	Title: Get AD Group Members
#	Author:			MTS-OFFICE\omenjivar
#	Originally created:	4/23/2020 - 08:03:29
#	Original path:		C:\Scripts\GroupMembersV1.ps1
#	Description: This Script get AD group members and expor them to CSV file under user profile Documents folder 
#	
# ========================================================

#region  ScriptForm  Designer

#region  Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion


#region  Event  Loop

function Main
{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($frm_GetADMember)
}


#endregion



#region Script Settings
#<ScriptSettings xmlns="http://tempuri.org/ScriptSettings.xsd">
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< frm_GetADMember >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$frm_GetADMember = New-Object System.Windows.Forms.Form
$frm_GetADMember.ClientSize = New-Object System.Drawing.Size(497, 437)
$frm_GetADMember.Text = "Get AD Group Members"
#~~< lbl_message >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$lbl_message = New-Object System.Windows.Forms.Label
$lbl_message.Location = New-Object System.Drawing.Point(12, 393)
$lbl_message.Size = New-Object System.Drawing.Size(349, 35)
$lbl_message.TabIndex = 6
$lbl_message.Text = "Members will be exported to your Documents Folder"
$lbl_message.TextAlign = [System.Drawing.ContentAlignment]::TopRight
$lbl_message.add_Click({Label1Click($lbl_message)})
#~~< btn_Export >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_Export = New-Object System.Windows.Forms.Button
$btn_Export.Location = New-Object System.Drawing.Point(367, 388)
$btn_Export.Size = New-Object System.Drawing.Size(118, 23)
$btn_Export.TabIndex = 5
$btn_Export.Text = "Export to CSV"
$btn_Export.UseVisualStyleBackColor = $true
$btn_Export.add_Click({Btn_ExportClick($btn_Export)})
#~~< lbl_Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$lbl_Group = New-Object System.Windows.Forms.Label
$lbl_Group.Location = New-Object System.Drawing.Point(55, 49)
$lbl_Group.Size = New-Object System.Drawing.Size(49, 25)
$lbl_Group.TabIndex = 4
$lbl_Group.Text = "Group"
$lbl_Group.add_Click({Label1Click($lbl_Group)})
#~~< DataGrid1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DataGrid1 = New-Object System.Windows.Forms.DataGrid
$DataGrid1.DataMember = ""
$DataGrid1.Location = New-Object System.Drawing.Point(55, 111)
$DataGrid1.Size = New-Object System.Drawing.Size(392, 261)
$DataGrid1.TabIndex = 3
$DataGrid1.HeaderForeColor = [System.Drawing.SystemColors]::ControlText
#~~< txt_Group >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$txt_Group = New-Object System.Windows.Forms.TextBox
$txt_Group.Location = New-Object System.Drawing.Point(124, 49)
$txt_Group.Size = New-Object System.Drawing.Size(206, 20)
$txt_Group.TabIndex = 2
$txt_Group.Text = ""
#~~< btn_GetMembers >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$btn_GetMembers = New-Object System.Windows.Forms.Button
$btn_GetMembers.Location = New-Object System.Drawing.Point(351, 47)
$btn_GetMembers.Size = New-Object System.Drawing.Size(96, 23)
$btn_GetMembers.TabIndex = 0
$btn_GetMembers.Text = "Get Members"
$btn_GetMembers.UseVisualStyleBackColor = $true
$btn_GetMembers.add_Click({Button1Click($btn_GetMembers)})
$frm_GetADMember.Controls.Add($lbl_message)
$frm_GetADMember.Controls.Add($btn_Export)
$frm_GetADMember.Controls.Add($lbl_Group)
$frm_GetADMember.Controls.Add($DataGrid1)
$frm_GetADMember.Controls.Add($txt_Group)
$frm_GetADMember.Controls.Add($btn_GetMembers)

#endregion

#region Custom Code
Import-Module "ActiveDirectory"
#endregion

#region Event Loop

function Main{
	[System.Windows.Forms.Application]::EnableVisualStyles()
	[System.Windows.Forms.Application]::Run($frm_GetADMember)
}

#endregion

#endregion

#region Event Handlers

function Btn_ExportClick( $object ){
		
		

		
	$path = $env:UserProfile+"\Documents\"+$group+"_members.csv"
	
	
	$Members | Export-Csv $path
	
	
		
	#  "Memebers has been xported as C:\temp\$filename"

$lbl_message.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
$lbl_message.Location.y.Equals(410)
$lbl_message.Location.x.Equals(124)
 	$lbl_message.Text= "Exported to $path "
	
	
}

function Label1Click( $object ){

}

function Button1Click( $object ){
	
	$Script:group = $txt_Group.Text
	$MembersArray = New-Object System.Collections.ArrayList 
	$Script:Members = Get-ADGroup $group | Get-ADGroupMember | Get-ADUser -Properties LastLogonDate | select Name, SamAccountName, SID, LastLogonDate
	$MembersArray.AddRange($Members)
	$DataGrid1.DataSource = $MembersArray
	$frm_GetADMember.Refresh
}

Main # This call must remain below all other event functions

#endregion
