<#

    A simple GUI for the PowerShell Book Generator

#>

Import-Module $PSScriptRoot\PowerShellBooks -Force


#region XAML window definition
# Right-click XAML and choose WPF/Edit... to edit WPF Design
# in your favorite WPF editing tool
$xaml = @'
<Window
   xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
   xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
   MinWidth="200"
   Width ="400"
   SizeToContent="Height"
   Title="PowerShell Book Generator"
   Topmost="false">
   <Grid Margin="10,10,10,10">
      <Grid.ColumnDefinitions>
         <ColumnDefinition Width="Auto"/>
         <ColumnDefinition Width="*"/>
         <ColumnDefinition Width="35"/>
      </Grid.ColumnDefinitions>
      <Grid.RowDefinitions>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="Auto"/>
         <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <TextBlock Grid.Column="0" Grid.Row="0" Grid.ColumnSpan="3" Margin="5">The PowerShell Book Generator</TextBlock>

        <Label Name="Link" Grid.Column="0" Grid.Row="4" Grid.ColumnSpan="3" Margin="5">www.andreasnick.com
            <Label.Style>
                <Style TargetType="Label">
                    <Setter Property="Foreground" Value="Blue" />
                    <Style.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Foreground" Value="Red" />
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </Label.Style>
        </Label>


        <TextBlock Grid.Column="0" Grid.Row="1" Margin="5">PowerShell Module</TextBlock>
      <TextBlock Grid.Column="0" Grid.Row="2" Margin="5">OutputFolder</TextBlock>
        <ComboBox Name="ComboModule" Grid.Column="1" Grid.Row="1" Margin="5"></ComboBox>
      
      <TextBox Name="TxtOutputFolder" Grid.Column="1" Grid.Row="2" Margin="5"></TextBox>

      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,10,0,0" Grid.Row="3" Grid.ColumnSpan="2">
        <Button Name="ButStart" MinWidth="80" Height="22" Margin="5">Start</Button>
        <Button Name="ButCancel" MinWidth="80" Height="22" Margin="5">Cancel</Button>
      </StackPanel>
        <Button Grid.Column="2" Grid.Row="2" Name="ButFilesearch" Width="25" Height="22" Margin="5">...</Button>
    </Grid>
</Window>
'@
#endregion

#region Code Behind
function Convert-XAMLtoWindow
{
  param
  (
    [Parameter(Mandatory=$true)]
    [string]
    $XAML
  )
  
  Add-Type -AssemblyName PresentationFramework
  
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  $result = [Windows.Markup.XAMLReader]::Load($reader)
  $reader.Close()
  $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
  while ($reader.Read())
  {
      $name=$reader.GetAttribute('Name')
      if (!$name) { $name=$reader.GetAttribute('x:Name') }
      if($name)
      {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
  }
  $reader.Close()
  $result
}

function Show-WPFWindow
{
  param
  (
    [Parameter(Mandatory=$true)]
    [Windows.Window]
    $Window
  )
  
  $result = $null
  $null = $window.Dispatcher.InvokeAsync{
    $result = $window.ShowDialog()
    Set-Variable -Name result -Value $result -Scope 1
  }.Wait()
  $result
}
#endregion Code Behind

#region Convert XAML to Window
$window = Convert-XAMLtoWindow -XAML $xaml 
#endregion

#region Define Event Handlers
# Right-Click XAML Text and choose WPF/Attach Events to
# add more handlers
$window.ButCancel.add_Click(
  {
    $window.DialogResult = $false
  }
)

$window.ButStart.add_Click(
  {
    #
    $window.ButCancel.IsEnabled = $false
    $window.ButStart.IsEnabled = $false
    if(Test-Path $Window.TxtOutputFolder.Text){
    
      $OutputDocument =$($Window.TxtOutputFolder.Text + "\PowerShell_With_" + $window.ComboModule.Text + '.pdf')
      $OutputDocument = $OutputDocument -replace '\\','\'
      
      Write-Output $OutputDocument
      
      New-PowerShellBook -Module $window.ComboModule.Text -OutputPdfDocument  $OutputDocument
      
      
    } else {
      Write-Error "The output path not exist!"
    }
    
    $window.ButCancel.IsEnabled = $true
    $window.ButStart.IsEnabled = $true
    
  }
)
$window.ButFilesearch.add_Click{
  # remove param() block if access to event information is not required
  [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.SelectedPath = $Window.TxtOutputFolder.Text
    $browse.Description = "Select a directory"
    
    if ($browse.ShowDialog() -eq "OK")
        {
        
           $Window.TxtOutputFolder.Text = $browse.SelectedPath
		
		
        } 
    
}
$window.Link.add_MouseDown{
  # remove param() block if access to event information is not required
  param
  (
    [Parameter(Mandatory)][Object]$sender,
    [Parameter(Mandatory)][Windows.Input.MouseButtonEventArgs]$e
  )
  
  start "https://www.andreasnick.com"
}



#endregion Event Handlers

#region Manipulate Window Content

Write-Verbose "Query PowerShell modules.... please wait" -Verbose
$window.ComboModule.ItemsSource = (Get-Command -Module * | select -Property Module -Unique).Module | sort
$window.ComboModule.SelectedIndex=1
$window.TxtOutputFolder.Text = $($env:Userprofile +'\Desktop\')


#endregion

# Show Window

<#
    while ($x -gt 0)
    {
    # Content
    }
#>

$result = Show-WPFWindow -Window $window

#region Process results
if ($result -eq $true)
{

}
else
{
  
}
#endregion Process results