<#
    .SYNOPSIS
        Various add on menus to extend what the PowerShell Integrated Scripting Environment (ISE) can already do.

    .DESCRIPTION
        Various add on menus to extend what the PowerShell Integrated Scripting Environment (ISE) can already do.

    .NOTES    
        Author: Boe Prox
        Version History:
            1.0 //Boe Prox - 07/11/2016
                - Initial update
#>

#region Helper Functions
Function Show-CommandParameter {
    [cmdletbinding()]
    Param (
        [parameter()]
        [string]$Command = $psISE.CurrentFile.Editor.SelectedText
    )
    $ParsedData = [system.management.automation.psparser]::Tokenize($Command,[ref]$null) | Select -First 1
    If ($ParsedData.Type -eq 'Command') {
        (Get-Command $Command | 
            Select -expand Parameters).GetEnumerator() | ForEach {
                $_.Value 
            } | Out-GridView -Title $Command   
    }
}
Function Show-StaticTypeData {
    [cmdletbinding()]
    Param (
        [parameter()]
        [String]$Command = $psISE.CurrentFile.Editor.SelectedText
    )
    $ParsedData = [system.management.automation.psparser]::Tokenize($Command,[ref]$null) | Select -First 1
    If ($ParsedData.Type -eq 'Type') {
        [type]($ParsedData.content -replace '^\[(.*)\]$','$1') | 
        Get-Member -Static | Out-GridView -Title $ParsedData.Content
    }
}
function Invoke-IndentFormat {	
	Param(
		$ScriptText
	) 
	$CurrentLevel = 0
	$ParseError = $null
	$Tokens = [System.Management.Automation.PSParser]::Tokenize($ScriptText,[ref]$ParseError)
	
	if($ParseError) { 
		$ParseError | Write-Error
		throw "The parser will not work properly with errors in the script, please modify based on the above errors and retry."
	}
	
	for($t = $Tokens.Count -1 ; $t -ge 1; $t--) {
		
		$Token = $Tokens[$t]
		$NextToken = $Tokens[$t-1]
		
		if ($Token.Type -eq 'GroupStart') { 
			$CurrentLevel-- 
		}  
		
		if ($NextToken.Type -eq 'NewLine' ) {
			# Grab Placeholders for the Space Between the New Line and the next token.
			$RemoveStart = $NextToken.Start + 2  
			$RemoveEnd = $Token.Start - $RemoveStart
			$IndentText = "    " * $CurrentLevel 
			$ScriptText = $ScriptText.Remove($RemoveStart,$RemoveEnd).Insert($RemoveStart,  $IndentText)
		}
		
		if ($token.Type -eq 'GroupEnd') { 
			$CurrentLevel++ 
		}     
	}
	
	$ScriptText            
}
#endregion Helper Functions

#region Show-CommandParameter - Ctrl+Alt+C
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Show-CommandParameter",{Show-CommandParameter},"Ctrl+Alt+C")
#endregion Show-CommandParameter - Ctrl+Alt+C

#region Show-StaticTypeData - Ctrl+Alt+T
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Show-StaticTypeData",{Show-StaticTypeData},"Ctrl+Alt+T") 
#endregion Show-StaticTypeData - Ctrl+Alt+T        

#region Save-SelectedSnippet - Alt+F8
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Save-SelectedSnippet",{
    $Text = $psISE.CurrentFile.Editor.SelectedText
    If (($Text -notmatch '^\s+$' -AND $Text.length -gt 0)) {
        Try {
            [void][Microsoft.VisualBasic.Interaction]
        } Catch {
            Add-Type –assemblyName Microsoft.VisualBasic
        } 
        $Name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Snippet Name", "Snippet Name")
        $Description = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Snippet Description", "Snippet Description")    
        If ($Name -and $Description) {	       	    
            New-IseSnippet -Description $Description -Title $Name -Text $Text
            Write-Host "New Snippet created!" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
},"Alt+F8")      
#endregion Save-SelectedSnippet - Alt+F8

#region Save-Snippet - Alt+F5
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Save-Snippet",{
    $Text = $psISE.CurrentFile.Editor.Text
    If (($Text -notmatch '^\s+$' -AND $Text.length -gt 0)) {
        Try {
            [void][Microsoft.VisualBasic.Interaction]
        } Catch {
            Add-Type –assemblyName Microsoft.VisualBasic
        }    
        $Name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Snippet Name", "Snippet Name")
        $Description = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Snippet Description", "Snippet Description")    
        If ($Name -and $Description -AND ($Text -notmatch '^\s+$' -AND $Text.length -gt 0)) {	    
            New-IseSnippet -Description $Description -Title $Name -Text $Text
            Write-Host "New Snippet created!" -ForegroundColor Yellow -BackgroundColor Black
        }
    }
},"Alt+F5")  
#endregion Save-Snippet - Alt+F5
 
#region Add-TODO - Alt+F2
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Add-TODO",{
    Try {
        [void][Microsoft.VisualBasic.Interaction]
    } Catch {
        Add-Type –assemblyName Microsoft.VisualBasic
    }    
    $TODO = [Microsoft.VisualBasic.Interaction]::InputBox("Enter TODO Statement", "#TODO")  
    If ($TODO) {	    
        $psISE.CurrentFile.Editor.InsertText("#TODO: $($TODO)")
    }
},"Alt+F2")            
#endregion Add-TODO - Alt+F2

#region Invoke-IndentFormat - Ctrl+Alt+B
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Invoke-IndentFormat",{
	$CurrentLine = $psISE.CurrentFile.Editor.CaretLine
    $CurrentColumn = $psISE.CurrentFile.Editor.CaretColumn
	$psISE.CurrentFile.Editor.Text = Invoke-IndentFormat $psISE.CurrentFile.Editor.Text
    $psISE.CurrentFile.Editor.SetCaretPosition($CurrentLine,$CurrentColumn)
},"Ctrl+Alt+B")
#endregion Invoke-IndentFormat - Ctrl+Alt+B

#region ISE Keyboard Shortcut - Alt+K
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("ISE Keyboard Shortcuts",{
    $Assembly = $psISE.GetType().Assembly
    $ResourceManager = New-Object System.Resources.ResourceManager -ArgumentList GuiStrings,$Assembly
    $ResourceSet = $ResourceManager.GetResourceSet((Get-Culture),$true,$true)
    $ResourceSet | Where-Object Name -match 'Shortcut\d?$|^F\d+Keyboard' | Sort-Object Value | Out-GridView
},"Alt+K") 
#endregion ISE Keyboard Shortcut - Alt+K

#region CopyTextToFile - Alt+C
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("CopyTextToFile",{
    $Editor = $psISE.CurrentFile.Editor
    $SelectedText = $Editor.SelectedText
    $File = $psISE.CurrentPowerShellTab.Files.Add() 
    $File.Editor.Text = $SelectedText 
    $File.Editor.SetCaretPosition(1,1)
},"Alt+C") 
#endregion CopyTextToFile - Alt+C

#region InsertBlockComment - Alt+M
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("InsertBlockComment",{
    $Editor = $psISE.CurrentFile.Editor
    $caretLine = $Editor.CaretLine
    $SelectedText = $Editor.SelectedText
    $NewText = "<#`n{0}`n#>" -f $SelectedText
    $Editor.InsertText($NewText)
    $Editor.SetCaretPosition($caretLine, 1)
},"Alt+M") 
#endregion InsertBlockComment - Alt+M

#region InsertLineComment - Alt+N
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("InsertLineComment",{
    $Editor = $psISE.CurrentFile.Editor
    If (-NOT $Editor.SelectedText) {
        $Editor.SelectCaretLine()
        $caretLine = $editor.CaretLine
        $SelectedText = $Editor.SelectedText
        $Editor.InsertText(("#{0}" -f $SelectedText))
    }
    Else {
        If ($Editor.SelectedText -match "`n") {
            $NewText = ($Editor.SelectedText -split "`n" | ForEach {
                If ($_) {
                    "#{0}" -f $_
                }
                Else {
                    $_
                }
            }) -join "`n"        
        }
        Else {
            $NewText = ($Editor.SelectedText -split "`r" | ForEach {
                If ($_) {
                    "#{0}" -f $_
                }
                Else {
                    $_
                }
            }) -join "`n"           
        }
        $Editor.InsertText($NewText)
        $Editor.SetCaretPosition($caretLine, 1)
    }
},"Alt+N") 
#endregion InsertLineComment - Alt+N

#region RemoveTrailingSpaces - Ctrl+Shift+Del
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("RemoveTrailingSpaces",{
    $Editor = $psISE.CurrentFile.Editor
    $caretLine = $editor.CaretLine
    $newText = foreach ( $line in $editor.Text.Split("`n") ) {
        $line -replace ("\s+$", "")
    }
    $editor.Text = $newText -join "`n"
    $editor.SetCaretPosition($caretLine, 1)
},"Ctrl+Shift+Del") 
#endregion RemoveTrailingSpaces - Ctrl+Shift+Del

#region RemoveBlankLines - Alt+Shift+Del
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("RemoveBlankLines",{
    $Editor = $psISE.CurrentFile.Editor
    $CaretLine = $editor.CaretLine
    If ($Editor.SelectedText) {
        $editor.InsertText(($Editor.SelectedText -replace '(?m)\s*$', ''))
    }
    Else {
        $Editor.Text = $Editor.Text -replace '(?m)\s*$', ''
    }
    $Editor.SetCaretPosition(1, 1)
},"Alt+Shift+Del") 
#endregion RemoveBlankLines - Alt+Shift+Del

#region InsertRegionBlock - Alt+R
[void]$psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("InsertRegionBlock",{
    $Editor = $psISE.CurrentFile.Editor
    Try {
        [void][Microsoft.VisualBasic.Interaction]
    } Catch {
        Add-Type –assemblyName Microsoft.VisualBasic
    } 
    $Name = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Region Information", "Region Information")
    $CaretLine = $Editor.CaretLine
    $SelectedText = $Editor.SelectedText
    $NewText = "#region {0}`n{1}`n#endregion {0}" -f $Name,$SelectedText
    $Editor.InsertText($NewText)
    $Editor.SetCaretPosition($CaretLine, 1)
},"Alt+R") 
#endregion InsertRegionBlock - Alt+R