<#
.DESCRIPTION
This is a simple GUI tool that will query Services, TaskSchedule on multiple computers and it will list it in table format. 
This tool has capability to stop a service, start a service, stop running task schedule, start task schedule. 
Each function can be accessed on each tab.

.AUTHOR
Leonard Sulit

#>

#Loading WindowsPresentationFramework
Add-Type -AssemblyName PresentationFramework
$VerbosePreference = 'Continue'

#Assigning xaml content to a powershell variable
$xmlPath = 'C:\users\lsulit\documents\Visual Studio 2017\Projects\HRSD-Tool\HRSD-Tool\MainWindow.xaml'
[xml]$Form = @"

<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="HRSD Tool" Height="487.714" Width="605.563" WindowStartupLocation="CenterScreen" ResizeMode="NoResize" Topmost="True">
    <Grid>
        <TabControl HorizontalAlignment="Left" Height="440" VerticalAlignment="Top" Width="586" Margin="4,2,0,0">
            <TabItem Header="Task Schedule">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="40*"/>
                        <ColumnDefinition Width="114*"/>
                        <ColumnDefinition Width="127*"/>
                        <ColumnDefinition Width="96*"/>
                        <ColumnDefinition Width="125*"/>
                        <ColumnDefinition Width="78*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="34*"/>
                        <RowDefinition Height="6*"/>
                        <RowDefinition Height="343*"/>
                        <RowDefinition Height="8*"/>
                        <RowDefinition Height="21*"/>
                    </Grid.RowDefinitions>
                    <Image Source="C:\users\lsulit\pictures\infor.png" Margin="0,0,0,0.6"/>
                    <TextBox x:Name="txt_taskCompName" Grid.Column="2" HorizontalAlignment="Left" Height="23" Margin="3.4,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Task Name: " Grid.Column="3" HorizontalAlignment="Left" Margin="4.2,5,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.66,0.084"/>
                    <TextBox x:Name="txt_TaskName" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="3.2,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Computer Name:" Grid.Column="1" HorizontalAlignment="Left" Margin="6,4,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_GetTask" Content="Get Task" Grid.Column="5" HorizontalAlignment="Left" Margin="3.4,8,0,0" VerticalAlignment="Top" Width="75" ClickMode="Press" IsDefault="True"/>
                    <DataGrid x:Name="dg_TaskSchedule" Grid.ColumnSpan="6" Grid.Row="2"/>
                    <Button x:Name="btn_RunTask" Content="Run Task" HorizontalAlignment="Left" Margin="9,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.ColumnSpan="2" Grid.RowSpan="2" ClickMode="Press"/>
                    <Button x:Name="btn_EndTask" Grid.ColumnSpan="2" Content="End Task" Grid.Column="1" HorizontalAlignment="Left" Margin="49,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.RowSpan="2" ClickMode="Press"/>
                </Grid>
            </TabItem>
            <TabItem Header="Services">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="40*"/>
                        <ColumnDefinition Width="114*"/>
                        <ColumnDefinition Width="127*"/>
                        <ColumnDefinition Width="96*"/>
                        <ColumnDefinition Width="125*"/>
                        <ColumnDefinition Width="78*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="34*"/>
                        <RowDefinition Height="6*"/>
                        <RowDefinition Height="343*"/>
                        <RowDefinition Height="8*"/>
                        <RowDefinition Height="21*"/>
                    </Grid.RowDefinitions>
                    <Image Source="C:\users\lsulit\pictures\infor.png" Margin="0,0,0,0.6"/>
                    <TextBox x:Name="txt_ComputerName" Grid.Column="2" HorizontalAlignment="Left" Height="23" Margin="3.4,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Service Name: " Grid.Column="3" HorizontalAlignment="Left" Margin="4.2,5,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.66,0.084"/>
                    <TextBox x:Name="txt_ServiceName" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="3.2,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Computer Name:" Grid.Column="1" HorizontalAlignment="Left" Margin="6,4,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_GetService" Content="Get Service" Grid.Column="5" HorizontalAlignment="Left" Margin="3.4,8,0,0" VerticalAlignment="Top" Width="75" ClickMode="Press" IsDefault="True"/>
                    <DataGrid x:Name="dg_Services" Grid.ColumnSpan="6" Grid.Row="2" CanUserSortColumns="True" CanUserReorderColumns="True" IsReadOnly="True" EnableColumnVirtualization="True" Margin="0,0,78,0.2">
                        <DataGrid.CurrentColumn>
                            <DataGridTemplateColumn SortMemberPath="Name" CanUserResize="False"/>
                        </DataGrid.CurrentColumn>
                    </DataGrid>
                    <Button x:Name="btn_StartService" Content="Start Service" HorizontalAlignment="Left" Margin="9,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.ColumnSpan="2" Grid.RowSpan="2" ClickMode="Press"/>
                    <Button x:Name="btn_StopService" Grid.ColumnSpan="2" Content="Stop Service" Grid.Column="1" HorizontalAlignment="Left" Margin="49,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.RowSpan="2" ClickMode="Press"/>
                    <CheckBox x:Name="chk_phmanlsulit01" Content="phmanlsulit01" Grid.Column="5" HorizontalAlignment="Left" Margin="0.4,5,0,0" Grid.Row="2" VerticalAlignment="Top" RenderTransformOrigin="-0.278,1.408" ClickMode="Press"/>
                    <CheckBox x:Name="chk_localhost" Content="localhost" Grid.Column="5" HorizontalAlignment="Left" Margin="1.4,21,0,0" Grid.Row="2" VerticalAlignment="Top" ClickMode="Press"/>
                </Grid>
            </TabItem>
            <TabItem Header="Application Pool">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="40*"/>
                        <ColumnDefinition Width="114*"/>
                        <ColumnDefinition Width="127*"/>
                        <ColumnDefinition Width="96*"/>
                        <ColumnDefinition Width="125*"/>
                        <ColumnDefinition Width="78*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="34*"/>
                        <RowDefinition Height="6*"/>
                        <RowDefinition Height="343*"/>
                        <RowDefinition Height="8*"/>
                        <RowDefinition Height="21*"/>
                    </Grid.RowDefinitions>
                    <Image Source="C:\users\lsulit\pictures\infor.png" Margin="0,0,0,0.6"/>
                    <TextBox x:Name="txt_AppCompName" Grid.Column="2" HorizontalAlignment="Left" Height="23" Margin="3.4,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Task Name: " Grid.Column="3" HorizontalAlignment="Left" Margin="4.2,5,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.66,0.084"/>
                    <TextBox x:Name="txt_AppName" Grid.Column="4" HorizontalAlignment="Left" Height="23" Margin="3.2,6,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Computer Name:" Grid.Column="1" HorizontalAlignment="Left" Margin="6,4,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_GetAppPool" Content="Get Task" Grid.Column="5" HorizontalAlignment="Left" Margin="3.4,8,0,0" VerticalAlignment="Top" Width="75" ClickMode="Press" IsDefault="True"/>
                    <DataGrid x:Name="dg_AppPool" Grid.ColumnSpan="6" Grid.Row="2"/>
                    <Button x:Name="btn_RecycleApp" Content="Recycle App" HorizontalAlignment="Left" Margin="9,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.ColumnSpan="2" Grid.RowSpan="2" ClickMode="Press"/>
                    <Button x:Name="btn_StopApp" Grid.ColumnSpan="2" Content="Stop App" Grid.Column="1" HorizontalAlignment="Left" Margin="49,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.RowSpan="2" ClickMode="Press"/>
                    <Button x:Name="btn_StartApp" Content="Start App" Grid.Column="2" HorizontalAlignment="Left" Margin="15.4,5.8,0,0" Grid.Row="3" VerticalAlignment="Top" Width="75" Grid.RowSpan="2"/>
                </Grid>
            </TabItem>
        </TabControl>

    </Grid>
</Window>

"@

#Assigning XamlReader to variable aswell loading a new object xml.xmlnodereader and feeding xaml content from $Form variable  as its argument
$Window = [Windows.Markup.XamlReader]::Load((New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $Form))

#region Global variables
#XAML Control Variables Map to Powershell Variables

#Button Variables
$Btn_StartService = $window.findname("btn_StartService")
$Btn_StopService = $window.findname("btn_StopService")
$Btn_DisableService = $window.findname("btn_DisableService")
$Btn_GetService = $window.FindName("btn_GetService")

#Checkbox Variables
$Chk_localhost = $window.FindName("chk_localhost")
$Chk_phmanlsulit01 = $window.FindName("chk_phmanlsulit01")

#DataGrid Variables
$Dg_Services = $window.FindName("dg_Services")

#TextBox Variables
$Txt_ComputerName = $window.FindName("txt_ComputerName")
$Txt_ComputerName.Focus()
$Txt_ServiceName = $window.FindName("txt_ServiceName")



#endregion

#region Custom Functions




#endregion

#GLOBAL
$checkedList = New-Object -TypeName System.Collections.ArrayList

#region Checkbox Click event localhost


$Chk_localhost.add_checked({

    Write-Verbose "$($this.name) State: Checked"
    $checkedList.Add([string]$Chk_localhost.Content)
    Write-Verbose "current value of the variable is: $checkedList"
})

$Chk_localhost.add_unchecked({
    
    Write-Verbose "$($this.name) State: Unchecked"
    $checkedList.remove([string]$Chk_localhost.Content)
    Write-Verbose "current value of the variable is: $checkedList"
})


#endregion Checkbox Click event localhost

#region Checkbox Click event phmanlsulit01

$Chk_phmanlsulit01.add_checked({

    Write-Verbose  "$($this.name) State: Checked"
    $checkedList.Add([string]$Chk_phmanlsulit01.Content)
    Write-Verbose  "current value of the variable is: $checkedList"
})

$Chk_phmanlsulit01.add_unchecked({
    
    Write-Verbose  "$($this.name) State: Unchecked"
    $checkedList.remove([string]$Chk_phmanlsulit01.Content)
    Write-Verbose  "current value of the variable is: $checkedList"
   
})
#endregion Checkbox Click event phmanlsulit01

#region Button Click event Get Service
$Btn_GetService.Add_Click({

    if($Txt_ServiceName.Text -eq "" -or $null -and $Txt_ComputerName.Text -eq "" -or $null){

        [System.Windows.MessageBox]::Show("Please enter a computer name and service name", "Result")

    }elseif($Txt_ServiceName.Text -eq "" -or $Txt_ServiceName.Text -eq $null) {

        [System.Windows.MessageBox]::Show("Please enter a service name", "Result")
        
    }elseif($Txt_ComputerName.Text -eq "" -or $Txt_ComputerName.Text -eq $null) {

        [System.Windows.MessageBox]::Show("Please enter a computer name and try again", "Result")

    }else {
    
        Write-Verbose "Creating object to store arraylist"
        $Service_Array = new-object -TypeName System.Collections.ArrayList

        #Creating array variable and split each string separated by comma.
        #$checkedList += $Txt_ComputerName.Text.Split(",")
        [array]$ServList = $Txt_ServiceName.Text.Split(",")

        #looping through each computer assigned to comp variable.
        foreach ($comp in $checkedList) {

            foreach($serv in $ServList) {

                [array]$ServiceList += Get-ServiceList -ServiceName $serv -ComputerName $comp  -Verbose   
            }

        }
    
        
        #Populating the Datagrid
        if ($ServiceList -ne $null) {

            $Service_Array.AddRange($ServiceList)
            $Dg_Services.ItemsSource = @($Service_Array)
    
        }else {

            $Dg_Services.ItemsSource = $null
            [System.Windows.MessageBox]::Show("Sorry there is no service found with the name: $Serv, please try again", "Result")
        }

    }

})
#endregion

#region Button Click event Stop Service
$Btn_StopService.Add_Click({
     
    [array]$SelectedServices = $Dg_Services.SelectedItems
    

    foreach($item in $SelectedServices) {

        $ServName = $item.Name
        $CompName =  $item.MachineName
        Stop-ServiceList -ComputerName $CompName -ServiceName $ServName
        Write-Verbose "Service: $ServName in $CompName has been successfully Stopped"

    }
    
    $Service_Array = new-object -TypeName System.Collections.ArrayList

    #Creating array variable and Manipulate string to split each string separated by comma.
    [array]$CompList = $Txt_ComputerName.Text.Split(",")
    [array]$ServList = $Txt_ServiceName.Text.Split(",")

    #looping through each computer assigned to comp variable.
    foreach ($comp in $CompList) {

        foreach($serv in $ServList) {

            [array]$ServiceList += Get-ServiceList -ServiceName $serv -ComputerName $comp  -Verbose 
        }

    }

    #Populating the Datagrid
    if ($ServiceList -ne $null) {

        $Service_Array.AddRange($ServiceList)
        $Dg_Services.ItemsSource = @($Service_Array)
    }
})
#endregion

#region Button Click event Start Service
$Btn_StartService.Add_Click({
    
    [array]$SelectedServices = $Dg_Services.SelectedItems
    

    foreach($item in $SelectedServices) {

        $ServName = $item.Name
        $CompName =  $item.MachineName

        Start-ServiceList -ComputerName $CompName -ServiceName $ServName -Verbose
        Write-Verbose "Service: $ServName in $CompName has been successfully Started"
        
    }
    

    $Service_Array = new-object -TypeName System.Collections.ArrayList

    #Creating array variable and Manipulate string to split each string separated by comma.
    [array]$CompList = $Txt_ComputerName.Text.Split(",")
    [array]$ServList = $Txt_ServiceName.Text.Split(",")

    #looping through each computer assigned to comp variable.
    foreach ($comp in $CompList) {

        foreach($serv in $ServList) {

            [array]$ServiceList += Get-ServiceList -ServiceName $serv -ComputerName $comp  -Verbose   
        }

    }

    #Populating the Datagrid
    if ($ServiceList -ne $null) {

        $Service_Array.AddRange($ServiceList)
        $Dg_Services.ItemsSource = @($Service_Array)
    }
})
#endregion

#Calling the window 
$Window.showdialog() | Out-Null