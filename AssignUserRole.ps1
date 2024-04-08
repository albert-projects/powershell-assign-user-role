
Add-Type -AssemblyName PresentationFramework

Import-Module ActiveDirectory

$dir = "."
$jsonfile = "AssignUserRole.json"

$weekday = @("Monday", "Tuesday", "Wednesday", "Thursday", "Friday")   

#get ou path
$User_Structure = Get-ADOrganizationalUnit -Filter * -Properties Name, CanonicalName, DistinguishedName -SearchBase 'OU=Users,DC=<your domain>,DC=com,DC=au' |  Where-Object {$_.Name -ne 'Users' } | select Name, CanonicalName, DistinguishedName

$PathObj = @()

$User_Structure| ForEach-Object {

    $Object = New-Object PSObject
    $Object | add-member Noteproperty Name $_.Name

    $Path = $_.CanonicalName -replace '<your domain>/Users/', ''
    $Object | add-member Noteproperty Path $Path

    $Object | add-member Noteproperty DistinguishedName $_.DistinguishedName

    $PathObj += $Object

}

#$PathObj | Sort-Object Path | Export-Csv "$dir\path.csv" -NoTypeInformation 

#get ad user
$User = Get-ADUser -Filter {(Enabled -eq "True")} -properties  DisplayName, sAMAccountName, Enabled  -SearchBase 'OU=Users,DC=<your domain>,DC=com,DC=au' | Select DisplayName, sAMAccountName, Enabled

$UserObj = @()

$User| ForEach-Object {

    $Object = New-Object PSObject
    $Object | add-member Noteproperty DisplayName $_.DisplayName
    $Object | add-member Noteproperty sAMAccountName $_.sAMAccountName

    $tempName = $_.DisplayName + " (" + $_.sAMAccountName + ")"
    $Object | add-member Noteproperty Name $tempName

    $UserObj += $Object

}

#$UserObj | Sort-Object DisplayName | Export-Csv "$dir\user.csv" -NoTypeInformation 

#get security group
$Grps = Get-ADGroup -Filter{GroupCategory -eq "security"} -properties Name, DistinguishedName  -SearchBase 'OU=Users,DC=<your domain>,DC=com,DC=au' | select Name, DistinguishedName

$GrpObj = @()

$Grps | ForEach-Object {

    $Object = New-Object PSObject
    $Object | add-member Noteproperty Name $_.Name
    $Object | add-member Noteproperty DistinguishedName $_.DistinguishedName
    $GrpObj += $Object

}

#$GrpObj | Sort-Object Name | Export-Csv "$dir\role.csv" -NoTypeInformation 


#Json file location 
$configfile = $dir + "\" + $jsonfile

#Timestamp
$Date = [datetime]::parseexact((Get-Date -Format "dd/MM/yyyy"), 'dd/MM/yyyy', $null )


function LoadJson{

    #$lblMsg.Content = ""

    $datGrid.Clear()
    $Form.Resources.Clear()
    $json = Get-Content $configfile -raw | ConvertFrom-Json
    #$datGrid.ItemsSource = $json.staff
    #$datGrid.Columns[4].Visibility = "Collapsed"
    #$datGrid.DataContext    

    $datGrid.ItemsSource = $json.staff
    #$datGrid.AddChild([pscustomobject]@{User='Albert';EffectiveDate='Monday';Role='Network Engineer';OU='IT_DEPT';EndDate='15/02/2024'})
    
    $Userdata = $UserObj | Select-Object -ExpandProperty Name -Unique | Sort
    #Write-Host $Userdata
    $Form.Resources.Add("User", $Userdata)
        
    $Grpdata = $GrpObj | Select-Object -ExpandProperty Name -Unique | Sort
    $Form.Resources.Add("Role", $Grpdata)

    $Pathdata = $PathObj | Select-Object -ExpandProperty Path -Unique | Sort
    $Form.Resources.Add("OU", $Pathdata)

    $Date = [datetime]::parseexact((Get-Date -Format "dd/MM/yyyy"), 'dd/MM/yyyy', $null )
    $Form.Resources.Add("EndDate", $Date)

    $Form.Resources.Add("EffectiveDate", $weekday)

    $Form.Resources.Add("Remove", "false")

}


function SaveJson{

    $GridItem = @()
    
    #Write-Host $datGrid.items.Count.ToString()
    #Write-Host $datGrid.Columns.Count.ToString()
    #MessageBox.Show($datGrid.SelectedCells[0].Value.ToString());

    $empty_flag = 0

    foreach ( $i in $datGrid.items ){
    
        if(($i.enddate.ToString() -eq "") -or ($i.effectivedate -eq "") -or ($i.user -eq "") -or ($i.role -eq "") -or ($i.ou -eq "")){
            $empty_flag++
        }
        if ($i.effectivedate.ToString() -eq ""){
           [System.Windows.Forms.MessageBox]::Show("EffectiveDate could not be empty","EffectiveDate could not be empty", "OK" , "Warning" )      
        }
        if ($i.user.ToString() -eq ""){
           [System.Windows.Forms.MessageBox]::Show("User could not be empty","User could not be empty", "OK" , "Warning" )      
        }
        if ($i.role.ToString() -eq ""){
           [System.Windows.Forms.MessageBox]::Show("Role could not be empty","Role could not be empty", "OK" , "Warning" )      
        }
        if ($i.ou.ToString() -eq ""){
           [System.Windows.Forms.MessageBox]::Show("OU could not be empty","OU could not be empty", "OK" , "Warning" )      
        }
        if ($i.enddate.ToString() -eq ""){
           [System.Windows.Forms.MessageBox]::Show("EndDate could not be empty","EndDate could not be empty", "OK" , "Warning" )       
        }



        $Object = New-Object PSObject
        
        #Write-Host $i.enddate
        #$tmpdate = ''
        $tmpdate = $i.enddate.ToString()

        #Write-Host $tmpdate.length 
        #Write-Host $tmpdate

        if ($tmpdate.length -ge 10){
            $tmpdate = $tmpdate.substring(0, 10) 
        }
        $tmpdate = $tmpdate.Trim()
        #$tmpdate = ([DateTime]($tmpdate)).ToString('dd/MM/yyyy')
        
        $Object | add-member Noteproperty enddate $tmpdate
        $Object | add-member Noteproperty effectivedate $i.effectivedate
        $Object | add-member Noteproperty user $i.user
        $Object | add-member Noteproperty role $i.role
        $Object | add-member Noteproperty ou $i.ou
        $Object | add-member Noteproperty remove $i.remove

        $tmpUserID = $UserObj | Where-Object {$_.Name -eq $i.user } | select sAMAccountName
        $Object | add-member Noteproperty userid $tmpUserID.sAMAccountName
        #Write-Host $tmpUserID

        $tmpOUPath = $PathObj | Where-Object {$_.Path -eq $i.ou } | select DistinguishedName
        $Object | add-member Noteproperty oupath $tmpOUPath.DistinguishedName

        $GridItem += $Object
        #Write-Host $i     
    }

    if($empty_flag -gt 0){      
        return
    }

    
    #Write-Host $GridItem

    #write to json
    #$UserList = $GridItem | Select-Object -Property @{Name="enddate";Expression={$_.enddate.ToString("dd/MM/yyyy")}}, effectivedate, user, role, ou, remove, userid, oupath  | Where-Object {$_.remove -notmatch "True" } | select enddate, effectivedate, user, role, ou, remove, userid, oupath
    $UserList = $GridItem | Select-Object -Property enddate, effectivedate, user, role, ou, remove, userid, oupath  | Where-Object {$_.remove -notmatch "True" } | select enddate, effectivedate, user, role, ou, remove, userid, oupath


    $json = New-Object -TypeName PSObject
    $json | Add-Member -MemberType NoteProperty -Name "Staff" -Value @($UserList)
    $json | ConvertTo-Json -depth 100 | Out-File $configfile

    $lblMsg.Content = "Records saved"

    #Write-Host $UserList

    LoadJson


}

function AddRecord{

    #$datGrid.AddChild([pscustomobject]@{EffectiveDate='Monday';User='';Role='';OU='';EndDate='';Remove='false'})
    $source = $datGrid.ItemsSource
    $source += [pscustomobject]@{EffectiveDate='';User='';Role='';OU='';EndDate='';Remove="false"}
    $datGrid.ItemsSource = $source


}



[void][System.Reflection.Assembly]::LoadWithPartialName('CasualStaffSchedule')
[xml]$XAML = @"

<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Casual Staff Schedule (Version 1.0)" Height="780" Width="1280">
    <Grid Margin="0,0,0,0">
        <Image Name="imgIcon" HorizontalAlignment="Left" Height="100" Margin="20,20,0,0" VerticalAlignment="Top" Width="100" Source="logo.png"/>
        <Label Name="lblTitel" Content="Casual Staff Schedule" HorizontalAlignment="Left" Margin="153,52,0,0" VerticalAlignment="Top" FontSize="20" FontWeight="Bold"/>
        <DataGrid Name="datGrid" AutoGenerateColumns="False" HorizontalAlignment="Left" Height="510" Margin="20,142,0,0" VerticalAlignment="Top" Width="1200">
         <DataGrid.Columns>

                 <DataGridTemplateColumn Header="EffectiveDate" Width="100">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ComboBox
                                SelectedItem="{Binding Path=EffectiveDate, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                ItemsSource="{DynamicResource EffectiveDate}"
                                Text="{Binding Path=EffectiveDate, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            </ComboBox>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                 </DataGridTemplateColumn>

                 <DataGridTemplateColumn Header="User" Width="200">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ComboBox
                                SelectedItem="{Binding Path=User, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                ItemsSource="{DynamicResource User}"
                                Text="{Binding Path=User, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            </ComboBox>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>

                 <DataGridTemplateColumn Header="Role" Width="300" >
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ComboBox
                                SelectedItem="{Binding Path=Role, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                ItemsSource="{DynamicResource Role}"                                
                                Text="{Binding Path=Role, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            </ComboBox>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>

                <DataGridTemplateColumn Header="OU" Width="350">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ComboBox
                                SelectedItem="{Binding Path=OU, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                ItemsSource="{DynamicResource OU}"                                
                                Text="{Binding Path=OU, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            </ComboBox>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>

                <!--  
                <DataGridTextColumn Header="EffectiveDate" Binding="{Binding Path=EffectiveDate}" Width="100"/>
                <DataGridTextColumn Header="User" Binding="{Binding Path=User}"/> 
                <DataGridTextColumn Header="Role" Binding="{Binding Path=Role}"/> 
                <DataGridTextColumn Header="OU" Binding="{Binding Path=OU}"/>
                <DataGridTextColumn Header="EndDate" Binding="{Binding Path=EndDate}" Width="200"/>
                -->

                <DataGridTemplateColumn Header="EndDate" Width="150">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding EndDate, StringFormat=dd/MM/yyyy, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                    <DataGridTemplateColumn.CellEditingTemplate>
                        <DataTemplate>
                            <DatePicker SelectedDate="{Binding EndDate, StringFormat=dd/MM/yyyy, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" />
                        </DataTemplate>
                    </DataGridTemplateColumn.CellEditingTemplate>
                </DataGridTemplateColumn>

                <DataGridTemplateColumn Header="Remove" Width="80">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox HorizontalAlignment="Center" VerticalAlignment="Stretch" 
                            IsChecked="{Binding Remove, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" >
                            </CheckBox>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>


            </DataGrid.Columns> 
        </DataGrid>
        <Button Name="btnLoad" Content="Reload" HorizontalAlignment="Left" Height="45" Margin="20,675,0,0" VerticalAlignment="Top" Width="140"/>
        <Button Name="btnAdd" Content="Add" HorizontalAlignment="Left" Height="45" Margin="200,675,0,0" VerticalAlignment="Top" Width="140"/>
        <Button Name="btnSave" Content="Save" HorizontalAlignment="Left" Height="45" Margin="380,675,0,0" VerticalAlignment="Top" Width="140"/>
        <Label Name="lblMsg" Content="" HorizontalAlignment="Left" Margin="547,685,0,0" VerticalAlignment="Top" Width="360" Foreground="Red"/>
    </Grid>
</Window>


"@
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader"; exit}

# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}


#$dataGridView = New-Object System.Windows.Forms.DataGridView
#$dataGridView.Size=New-Object System.Drawing.Size(960,510)



#Assign event
$btnLoad.Add_Click({
    #$lblTitel.Content = "test"
    $lblMsg.Content = ""

    $reload = [System.Windows.Forms.MessageBox]::Show("Reload user role?","Reload user role", "OKCancel" , "Information" )

    if ($reload -eq 'OK') { 
        LoadJson
    }   

})


$btnSave.Add_Click({

    $save = [System.Windows.Forms.MessageBox]::Show("Save user role?","Save user role", "OKCancel" , "Information" )

    if ($save -eq 'OK') { 
        SaveJson
    }   

})

$btnAdd.Add_Click({

        AddRecord
 
})


#init load json
LoadJson

#Show Form
$Form.ShowDialog() | out-null

