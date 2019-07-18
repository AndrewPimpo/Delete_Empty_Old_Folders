### PURPOSE ###
# Find folders in the directory selected which are empty and 1 year old
# or which are just one of these things, and to consecutively delete them.
## - -
##### Functions #####
$directory = $null

#Select your directory
function select-Dir(){
    
    #Save dir path
    [string]$path = $null
    
    #Load file browser
	$openDir = new-object System.Windows.Forms.FolderBrowserDialog
	$openDir.RootFolder = "MyComputer"
	$openDir.Description = "Select the directory for analysis:"	

	#Show file browser
	$Show = $openDir.ShowDialog()
    
	#Use selected folder\drive as directory
	if($Show -eq "OK")
	{
		$path = $openDir.SelectedPath
    }
    else{
        $path = "Cancel"
    }

    return $path
}

#Pick the folder kinds
function filter-Processing(){

    #Load the assembly and the form
    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form

    #Form Label
    $l_Desc = New-Object System.Windows.Forms.Label
    $l_Desc.AutoSize = $true
    $l_Desc.Location = New-Object System.Drawing.Size(10,20)
    $l_Desc.Text = "RULE ALTERATIONS - Check boxes to alter filter rules so that the program:"
    $Form.Controls.Add($l_Desc)

    $Form.Width = ($l_Desc.Width + 40)

    #Form checkbox1
    $n_Empty = New-Object System.Windows.Forms.CheckBox
    $n_Empty.Location = New-Object System.Drawing.Size(30,40) 
    $n_Empty.Size = New-Object System.Drawing.Size(200,30)
    $n_Empty.Text = "Deletes folders that contain files."
    $Form.Controls.Add($n_Empty)

    #Form checkbox2
    $n_Old = New-Object System.Windows.Forms.CheckBox
    $n_Old.Location = New-Object System.Drawing.Size(30,70) 
    $n_Old.Size = New-Object System.Drawing.Size(200,30)
    $n_Old.Text = "Deletes folders < one year old."
    $Form.Controls.Add($n_Old)

    #Form button1
    $b_Set = New-Object System.Windows.Forms.Button
    $b_Set.Text = "Set"
    $b_Set.Location = New-Object System.Drawing.Size(100,125)
    $b_Set.Size = New-Object System.Drawing.Size(80,30)
    $b_Set.Add_Click({$Form.Close()})
    $b_Set.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.Controls.Add($b_Set)
    
    #Form button2
    $b_Cancel = New-Object System.Windows.Forms.Button
    $b_Cancel.Text = "Cancel"
    $b_Cancel.Location = New-Object System.Drawing.Size(200,125)
    $b_Cancel.Size = New-Object System.Drawing.Size(80,30)
    $b_Cancel.Add_Click({$Form.Close()})
    $b_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $Form.Controls.Add($b_Cancel)

    $Form.Height = ($b_Cancel.Location.Y + 80)

    $Show = $Form.ShowDialog()

    if($Form.DialogResult -eq "OK"){

        #Check if checkboxes were checked (hahaha...)
        if(($n_Empty.CheckState -eq $true) -and ($n_Old.CheckState -eq $true)){
            return 'a'
        }elseif($n_Empty.CheckState -eq $true){
            return 'b'
        }elseif($n_Old.CheckState -eq $true){
            return 'c'
        }else{
            return 'd'
        }

    } else{
        return $null
    }
}

#Get old, empty folders
function find-Old-Empty([string]$path, [char]$status){
    $run = $null

    switch($status){
        'a'{
            $P = Get-ChildItem -path $path -Recurse | ?{($_.PSIsContainer)}
        }
        'b'{
            $P = Get-ChildItem -path $path -Recurse | ?{($_.PSIsContainer) -and ($_.GetFiles().Count -eq 0) -and ($_.GetDirectories().Count -eq 0)}
        }
        'c'{
            $P = Get-ChildItem -path $path -Recurse | ?{($_.PSIsContainer) -and ($_.CreationTime -lt (Get-Date).AddDays(-365))}
        }
        'd'{
            $P = Get-ChildItem -path $path -Recurse | ?{($_.PSIsContainer) -and ($_.GetFiles().Count -eq 0) -and ($_.GetDirectories().Count -eq 0) -and ($_.CreationTime -lt (Get-Date).AddDays(-365))}
        }
        default{
            exit
        }
    }

    #Verify that this is okay to delete
    $input = [System.Windows.MessageBox]::Show('NOTICE: This program will delete {0} folders. Continue?' -f ($P.Count),'Warning','OKCancel','Error')
    switch ($input){
        'OK'{
                $run = 'Y'
            }
        'Cancel'{
                $run = 'N'
            }
    }

    #Check action
    if($run -eq 'Y'){

        #Delete
        foreach($col in $P){
            Remove-Item -path $col.FullName
        }
    }

    return $run
}

#Send confirmation
function confirm-Result([char]$cmplt){
    
    #Check if find-Old-Empty deleted or not
    if($cmplt -eq 'Y'){
        [System.Windows.MessageBox]::Show('Folders deleted.','Complete','OK')
    }else{
        [System.Windows.MessageBox]::Show('Deletion Cancelled.','Cancelled','OK')
    }
}

##### Enter Program ######
$directory = select-Dir

$status = filter-Processing

$cmplt = find-Old-Empty -path $directory -status $status

confirm-Result -cmplt $cmplt

#

#

#

#

#

#

#