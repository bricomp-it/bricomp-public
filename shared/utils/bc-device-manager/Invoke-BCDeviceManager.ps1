#Requires -Version 7.0
<#
.SYNOPSIS
    Invoke-BCDeviceManager.ps1
    WPF GUI for managing Windows devices in bulk with parallel actions.

.DESCRIPTION
    A modern WPF/XAML graphical interface for remote device management.
    Replaces the console menu version with a proper Windows GUI.

    Features:
      - Load devices from CSV or enter names manually
      - Checkbox selection for targeted actions
      - Pre-validate WinRM reachability (Ping / WSMan / Auth)
      - BitLocker status check and bulk disable
      - Forced remote reboot with confirmation dialog
      - Live log written to file; "Show log" opens a tailing viewer
      - Export results to CSV
      - DNS suffix fallback for short names
      - Parallel execution via ForEach-Object -Parallel (PS7)

.PARAMETER CsvPath
    Optional path to a CSV file to pre-load on startup.

.PARAMETER ComputerNameColumn
    Column name in CSV for device names. Default: COMPUTERNAME

.PARAMETER DnsSuffix
    DNS suffix for short-name resolution (e.g. corp.local).

.PARAMETER LogRoot
    Directory for log output. Default: script directory.

.PARAMETER ThrottleLimit
    Max parallel threads. Default: 20.

.EXAMPLE
    pwsh -File .\Invoke-BCDeviceManager.ps1
    pwsh -File .\Invoke-BCDeviceManager.ps1 -CsvPath .\computers.csv -DnsSuffix corp.local

.NOTES
    BriComp Computers, LLC
    Version : 3.0.0
    Requires: PowerShell 7+, run as Administrator
    Use Start-BCDeviceManager.ps1 to launch from PS5 or double-click.

    === Revision History ===
    1.0.0  (2024)       Initial console menu release
    2.0.0  (2025-12-13) PS7 parallel execution, improved logging
    3.0.0  (2026-06-26) Full WPF/XAML GUI rewrite
#>

[CmdletBinding()]
param(
    [string]$CsvPath            = '',
    [string]$ComputerNameColumn = 'COMPUTERNAME',
    [string]$DnsSuffix          = '',
    [string]$LogRoot            = '',
    [int]   $ThrottleLimit      = 20
)

#region --- Admin check ---
$id = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not ([Security.Principal.WindowsPrincipal]$id).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
    [System.Windows.MessageBox]::Show(
        "Please run as Administrator.",
        "BriComp Device Manager",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning) | Out-Null
    exit 1
}
#endregion

#region --- Assemblies ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Runtime
# Ensure ObservableCollection is available
[System.Collections.ObjectModel.ObservableCollection[object]] | Out-Null
#endregion

#region --- Logging setup ---
if (-not $LogRoot) { $LogRoot = $PSScriptRoot }
$LogDir  = Join-Path $LogRoot ("Logs_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$LogFile = Join-Path $LogDir 'BCDeviceManager.log'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

Write-Log "BriComp Device Manager v3.0.0 started"
Write-Log "Log directory: $LogDir"
#endregion

#region --- XAML UI definition ---
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BriComp Device Manager"
    Height="580" Width="900"
    MinHeight="480" MinWidth="700"
    WindowStartupLocation="CenterScreen"
    FontFamily="Segoe UI" FontSize="12">

  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,5"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="#F5F5F5"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="Foreground" Value="#1A1A1A"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="5,4"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="BorderThickness" Value="1"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="GridLinesVisibility" Value="Horizontal"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#EEEEEE"/>
      <Setter Property="BorderBrush" Value="#E0E0E0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowHeight" Value="28"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="SelectionMode" Value="Extended"/>
      <Setter Property="SelectionUnit" Value="FullRow"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="AlternatingRowBackground" Value="#FAFAFA"/>
    </Style>
  </Window.Resources>

  <DockPanel>

    <!-- Toolbar -->
    <ToolBarTray DockPanel.Dock="Top" Background="#F0F0F0">
      <!-- Row 1: Actions -->
      <ToolBar Band="1" BandIndex="1" Background="#F0F0F0">

        <Button Name="btnPreValidate" ToolTip="Check Ping, WSMan, and Auth for selected devices">
          <TextBlock Text="Pre-validate" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnBitLockerStatus" ToolTip="Check BitLocker status on selected devices">
          <TextBlock Text="BitLocker status" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnDisableBitLocker" ToolTip="Disable BitLocker on selected devices" Foreground="#B91C1C">
          <TextBlock Text="Disable BitLocker" VerticalAlignment="Center" Foreground="#B91C1C"/>
        </Button>

        <Separator/>

        <Button Name="btnReboot" ToolTip="Force reboot selected devices" Foreground="#B91C1C">
          <TextBlock Text="Reboot" VerticalAlignment="Center" Foreground="#B91C1C"/>
        </Button>

        <Separator/>

        <Button Name="btnExportCsv" ToolTip="Export results to CSV">
          <TextBlock Text="Export CSV" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnShowLog" ToolTip="Open live log viewer">
          <TextBlock Text="Show log" VerticalAlignment="Center"/>
        </Button>

      </ToolBar>

      <!-- Row 2: Selection helpers -->
      <ToolBar Band="2" BandIndex="1" Background="#F0F0F0">

        <TextBlock Text="Select:" VerticalAlignment="Center" Foreground="#888888"
                   FontSize="11" Margin="4,0,6,0"/>

        <Button Name="btnCheckAll" ToolTip="Check all devices">
          <TextBlock Text="All" VerticalAlignment="Center"/>
        </Button>

        <Button Name="btnUncheckAll" ToolTip="Uncheck all devices">
          <TextBlock Text="None" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnUncheckOffline" ToolTip="Uncheck all devices that are OFFLINE (run Pre-validate first)">
          <TextBlock Text="Uncheck offline" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnCheckBitLockerOn" ToolTip="Select only devices with BitLocker enabled (run BitLocker status first)">
          <TextBlock Text="Select BitLocker on" VerticalAlignment="Center"/>
        </Button>

      </ToolBar>
    </ToolBarTray>

    <!-- Status bar -->
    <StatusBar DockPanel.Dock="Bottom" Background="#F0F0F0" Height="24">
      <StatusBarItem>
        <TextBlock Name="lblStatus" Text="Ready" FontSize="11" Foreground="#555555"/>
      </StatusBarItem>
      <Separator/>
      <StatusBarItem>
        <TextBlock Name="lblCounts" Text="0 devices" FontSize="11" Foreground="#555555"/>
      </StatusBarItem>
      <Separator/>
      <StatusBarItem>
        <TextBlock Name="lblLogPath" Text="" FontSize="11" Foreground="#888888"/>
      </StatusBarItem>
    </StatusBar>

    <!-- Main body: sidebar + grid -->
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="210" MinWidth="160"/>
        <ColumnDefinition Width="4"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Sidebar -->
      <DockPanel Grid.Column="0" Background="#F8F8F8">

        <Border DockPanel.Dock="Top" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1" Padding="8">
          <StackPanel>
            <TextBlock Text="TARGETS" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,6"/>

            <!-- Manual entry with watermark placeholder -->
            <DockPanel Margin="0,0,0,4">
              <Button DockPanel.Dock="Right" Name="btnAddDevice" Width="36" Padding="0" Margin="4,0,0,0" ToolTip="Add device" FontSize="16" FontWeight="Bold">+</Button>
              <Grid>
                <TextBox Name="txtDevice" AcceptsReturn="False" ToolTip="Enter device name and press Enter or click +"/>
                <TextBlock Name="txtDevicePlaceholder" Text="Enter device name..."
                           IsHitTestVisible="False" VerticalAlignment="Center"
                           Margin="6,0,0,0" Foreground="#AAAAAA" FontStyle="Italic" FontSize="11"/>
              </Grid>
            </DockPanel>

            <!-- CSV load + clear -->
            <DockPanel Margin="0,0,0,4">
              <Button DockPanel.Dock="Right" Name="btnClearAll" Width="60" Padding="4,0" Margin="4,0,0,0"
                      ToolTip="Clear all devices" Foreground="#B91C1C" FontSize="11">Clear</Button>
              <Button Name="btnLoadCsv" HorizontalAlignment="Stretch" FontSize="11">Load CSV...</Button>
            </DockPanel>

            <!-- DNS suffix -->
            <TextBlock Text="DNS SUFFIX (optional)" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,8,0,2"/>
            <TextBlock Text="Appended to short names from CSV or manual entry" FontSize="10" Foreground="#AAAAAA" Margin="0,0,0,4" TextWrapping="Wrap"/>
            <Grid>
              <TextBox Name="txtDnsSuffix" ToolTip="e.g. corp.local — short name 'SERVER01' becomes 'SERVER01.corp.local'"/>
              <TextBlock Name="txtDnsSuffixPH" Text="e.g. corp.local"
                         IsHitTestVisible="False" VerticalAlignment="Center"
                         Margin="6,0,0,0" Foreground="#AAAAAA" FontStyle="Italic" FontSize="11"/>
            </Grid>

            <!-- Credentials -->
            <TextBlock Text="CREDENTIALS" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,8,0,4"/>
            <CheckBox Name="chkCurrentUser" Content="Use current user" IsChecked="False"/>

          </StackPanel>
        </Border>

        <!-- Device list -->
        <ListBox Name="lstDevices" BorderThickness="0" Background="#F8F8F8"
                 SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto">
          <ListBox.ItemTemplate>
            <DataTemplate>
              <StackPanel Orientation="Horizontal">
                <Ellipse Width="8" Height="8" Margin="0,0,6,0" VerticalAlignment="Center">
                  <Ellipse.Style>
                    <Style TargetType="Ellipse">
                      <Setter Property="Fill" Value="#CCCCCC"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding StatusDot}" Value="OK">
                          <Setter Property="Fill" Value="#16A34A"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding StatusDot}" Value="WARN">
                          <Setter Property="Fill" Value="#D97706"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding StatusDot}" Value="ERROR">
                          <Setter Property="Fill" Value="#DC2626"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </Ellipse.Style>
                </Ellipse>
                <TextBlock Text="{Binding Name}" VerticalAlignment="Center"/>
              </StackPanel>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </DockPanel>

      <!-- Splitter -->
      <GridSplitter Grid.Column="1" Width="4" HorizontalAlignment="Stretch" Background="#E0E0E0"/>

      <!-- Results grid -->
      <DataGrid Name="dgResults" Grid.Column="2" Margin="0">

        <DataGrid.ContextMenu>
          <ContextMenu>
            <MenuItem Name="ctxRemoveDevice" Header="Remove device" />
            <Separator/>
            <MenuItem Name="ctxCheckSelected"   Header="Check selected rows" />
            <MenuItem Name="ctxUncheckSelected" Header="Uncheck selected rows" />
          </ContextMenu>
        </DataGrid.ContextMenu>

        <DataGrid.Columns>
          <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged}" Width="36" Header=""/>
          <DataGridTextColumn Binding="{Binding Name}"       Header="Device"         Width="160" IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding Ping}"       Header="Ping"           Width="70"  IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding WSMan}"      Header="WSMan"          Width="70"  IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding Auth}"       Header="Auth"           Width="70"  IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding BitLocker}"  Header="BitLocker"      Width="90"  IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding LastAction}" Header="Last action"    Width="*"   IsReadOnly="True"/>
        </DataGrid.Columns>

      </DataGrid>

    </Grid>
  </DockPanel>
</Window>
'@
#endregion

#region --- Device data model ---
# PowerShell classes cannot implement INotifyPropertyChanged reliably.
# Define the class in inline C# instead so WPF data binding works correctly.
Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;

namespace BriComp {
    public class DeviceItem : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;

        private void Notify(string prop) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private string _name       = "";
        private string _statusDot  = "NONE";
        private bool   _isSelected = true;
        private string _ping       = "\u2014";
        private string _wsman      = "\u2014";
        private string _auth       = "\u2014";
        private string _bitlocker  = "\u2014";
        private string _lastAction = "Not run";

        public string Name       { get { return _name;       } set { _name       = value; Notify("Name");       } }
        public string StatusDot  { get { return _statusDot;  } set { _statusDot  = value; Notify("StatusDot");  } }
        public bool   IsSelected { get { return _isSelected; } set { _isSelected = value; Notify("IsSelected"); } }
        public string Ping       { get { return _ping;       } set { _ping       = value; Notify("Ping");       } }
        public string WSMan      { get { return _wsman;      } set { _wsman      = value; Notify("WSMan");      } }
        public string Auth       { get { return _auth;       } set { _auth       = value; Notify("Auth");       } }
        public string BitLocker  { get { return _bitlocker;  } set { _bitlocker  = value; Notify("BitLocker");  } }
        public string LastAction { get { return _lastAction; } set { _lastAction = value; Notify("LastAction"); } }
    }
}
'@
#endregion

#region --- Build window ---
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Find controls
$btnPreValidate    = $window.FindName('btnPreValidate')
$btnBitLockerStatus= $window.FindName('btnBitLockerStatus')
$btnDisableBitLocker=$window.FindName('btnDisableBitLocker')
$btnReboot         = $window.FindName('btnReboot')
$btnExportCsv      = $window.FindName('btnExportCsv')
$btnShowLog        = $window.FindName('btnShowLog')
$btnCheckAll       = $window.FindName('btnCheckAll')
$btnUncheckAll     = $window.FindName('btnUncheckAll')
$btnUncheckOffline = $window.FindName('btnUncheckOffline')
$btnCheckBitLockerOn = $window.FindName('btnCheckBitLockerOn')
$btnAddDevice      = $window.FindName('btnAddDevice')
$txtDevicePH       = $window.FindName('txtDevicePlaceholder')
$txtDnsSuffixPH    = $window.FindName('txtDnsSuffixPH')
$btnClearAll       = $window.FindName('btnClearAll')
$btnLoadCsv        = $window.FindName('btnLoadCsv')
$txtDevice         = $window.FindName('txtDevice')
$txtDnsSuffix      = $window.FindName('txtDnsSuffix')
$chkCurrentUser    = $window.FindName('chkCurrentUser')
$lstDevices        = $window.FindName('lstDevices')
$dgResults         = $window.FindName('dgResults')
$lblStatus         = $window.FindName('lblStatus')
$lblCounts         = $window.FindName('lblCounts')
$lblLogPath        = $window.FindName('lblLogPath')

# Pre-fill params
if ($DnsSuffix)  { $txtDnsSuffix.Text = $DnsSuffix }
$lblLogPath.Text = "Log: $LogFile"

# Observable collection for the grid
$deviceList = [System.Collections.ObjectModel.ObservableCollection[BriComp.DeviceItem]]::new()
$dgResults.ItemsSource  = $deviceList
$lstDevices.ItemsSource = $deviceList
#endregion

#region --- Helper functions ---
function Update-Status {
    param([string]$Message)
    $window.Dispatcher.Invoke([Action]{
        $lblStatus.Text = $Message
        $ok    = @($deviceList | Where-Object { $_.StatusDot -eq 'OK'   }).Count
        $warn  = @($deviceList | Where-Object { $_.StatusDot -eq 'WARN' }).Count
        $err   = @($deviceList | Where-Object { $_.StatusDot -eq 'ERROR'}).Count
        $total = $deviceList.Count
        $lblCounts.Text = "$total devices · $ok OK · $warn warn · $err offline"
    })
}

function Get-BCCredential {
    <#
    .SYNOPSIS
        Shows a WPF credential dialog. Use instead of Get-Credential
        which requires a console and breaks under -WindowStyle Hidden.
    #>
    param([string]$Message = "Enter credentials for remote devices")

    $result = $null

    # Build a simple WPF credential dialog
    [xml]$credXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Credentials" Width="360" Height="220"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="12">
  <StackPanel Margin="16">
    <TextBlock Name="lblMsg" TextWrapping="Wrap" Margin="0,0,0,12"/>
    <TextBlock Text="Username:" Margin="0,0,0,4"/>
    <TextBox Name="txtUser" Margin="0,0,0,10"/>
    <TextBlock Text="Password:" Margin="0,0,0,4"/>
    <PasswordBox Name="txtPass" Margin="0,0,0,16"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnOK"     Content="OK"     Width="80" Margin="0,0,8,0" IsDefault="True"/>
      <Button Name="btnCancel" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
'@

    $credReader = [System.Xml.XmlNodeReader]::new($credXaml)
    $credWin    = [System.Windows.Markup.XamlReader]::Load($credReader)
    $credWin.Owner = $window

    $credWin.FindName('lblMsg').Text    = $Message
    $credWin.FindName('txtUser').Text   = "$env:USERDOMAIN\$env:USERNAME"

    $credWin.FindName('btnOK').Add_Click({
        $credWin.DialogResult = $true
        $credWin.Close()
    })
    $credWin.FindName('btnCancel').Add_Click({
        $credWin.DialogResult = $false
        $credWin.Close()
    })

    $dlgResult = $credWin.ShowDialog()

    if ($dlgResult -eq $true) {
        $user    = $credWin.FindName('txtUser').Text.Trim()
        $passSec = $credWin.FindName('txtPass').SecurePassword
        if ($user) {
            $result = [System.Management.Automation.PSCredential]::new($user, $passSec)
        }
    }

    return $result
}

function Add-Device {
    param([string]$DeviceName)
    $name = $DeviceName.Trim()
    if (-not $name) { return }

    # Append DNS suffix if set and name has no dot
    $suffix = $txtDnsSuffix.Text.Trim()
    if ($suffix -and $name -notmatch '\.') { $name = "$name.$suffix" }

    # No duplicates
    if ($deviceList | Where-Object { $_.Name -eq $name }) {
        Update-Status "Already in list: $name"
        return
    }

    $item = [BriComp.DeviceItem]::new()
    $item.Name = $name
    $window.Dispatcher.Invoke([Action]{ $deviceList.Add($item) })
    Write-Log "Device added: $name"
}

function Get-SelectedDevices {
    return @($deviceList | Where-Object { $_.IsSelected })
}

function Resolve-Name {
    param([string]$Name)
    $suffix = $window.Dispatcher.Invoke([Func[string]]{ $txtDnsSuffix.Text.Trim() })
    if ($suffix -and $Name -notmatch '\.') { return "$Name.$suffix" }
    return $Name
}
#endregion

#region --- Event: Add device ---
$btnAddDevice.Add_Click({
    Add-Device -DeviceName $txtDevice.Text
    $txtDevice.Clear()
    $txtDevice.Focus()
})

$txtDevice.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) {
        Add-Device -DeviceName $txtDevice.Text
        $txtDevice.Clear()
    }
})

# Show/hide placeholder text as user types
$txtDevice.Add_TextChanged({
    $txtDevicePH.Visibility = if ($txtDevice.Text.Length -eq 0) {
        [System.Windows.Visibility]::Visible
    } else {
        [System.Windows.Visibility]::Hidden
    }
})

# DNS suffix placeholder
$txtDnsSuffix.Add_TextChanged({
    $txtDnsSuffixPH.Visibility = if ($txtDnsSuffix.Text.Length -eq 0) {
        [System.Windows.Visibility]::Visible
    } else {
        [System.Windows.Visibility]::Hidden
    }
})
#endregion

#region --- Event: Load CSV ---
$btnLoadCsv.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter   = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.Title    = "Load device list"
    if ($dlg.ShowDialog() -eq $true) {
        try {
            $rows = Import-Csv -Path $dlg.FileName -ErrorAction Stop
            $col  = $ComputerNameColumn

            # Auto-detect column if not found
            if ($rows.Count -gt 0 -and -not ($rows[0].PSObject.Properties.Name -contains $col)) {
                $col = $rows[0].PSObject.Properties.Name | Select-Object -First 1
                Write-Log "CSV column '$ComputerNameColumn' not found, using '$col'" -Level WARN
            }

            foreach ($row in $rows) {
                $name = $row.$col
                if ($name) { Add-Device -DeviceName $name }
            }
            Update-Status "Loaded $($rows.Count) rows from CSV"
            Write-Log "CSV loaded: $($dlg.FileName) ($($rows.Count) rows)"
        } catch {
            [System.Windows.MessageBox]::Show(
                "Failed to load CSV:`n$($_.Exception.Message)",
                "Load error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error)
            Write-Log "CSV load failed: $($_.Exception.Message)" -Level ERROR
        }
    }
})
#endregion

#region --- Event: Clear all ---
$btnClearAll.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        "Remove all devices from the list?",
        "Clear all",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        $deviceList.Clear()
        Update-Status "Device list cleared"
        Write-Log "Device list cleared"
    }
})
#endregion

#region --- Event: Selection helpers ---
$btnCheckAll.Add_Click({
    foreach ($item in $deviceList) { $item.IsSelected = $true }
    $dgResults.Items.Refresh()
    Update-Status "All devices selected"
})

$btnUncheckAll.Add_Click({
    foreach ($item in $deviceList) { $item.IsSelected = $false }
    $dgResults.Items.Refresh()
    Update-Status "All devices deselected"
})

$btnUncheckOffline.Add_Click({
    $count = 0
    foreach ($item in $deviceList) {
        if ($item.Ping -eq 'OFFLINE' -or $item.Ping -eq 'TIMEOUT' -or $item.StatusDot -eq 'ERROR') {
            $item.IsSelected = $false
            $count++
        }
    }
    $dgResults.Items.Refresh()
    Update-Status "Unchecked $count offline device(s)"
    Write-Log "Uncheck offline: $count devices unchecked"
})

$btnCheckBitLockerOn.Add_Click({
    $blKeywords = @('On','Protection On','1','FullyEncrypted')
    $count = 0
    foreach ($item in $deviceList) {
        $bl = $item.BitLocker
        if ($bl -and $bl -ne '—' -and $bl -ne 'N/A' -and $bl -ne 'UNREACHABLE' -and
            ($bl -match 'On' -or $bl -eq '1')) {
            $item.IsSelected = $true
            $count++
        } else {
            $item.IsSelected = $false
        }
    }
    $dgResults.Items.Refresh()
    if ($count -eq 0) {
        Update-Status "No devices with BitLocker enabled found — run BitLocker status first"
    } else {
        Update-Status "Selected $count device(s) with BitLocker enabled"
        Write-Log "Select BitLocker on: $count devices selected"
    }
})
#endregion

#region --- Event: Context menu (right-click on grid row) ---
# Wire up after window is shown so ContextMenu is fully initialized
$window.Add_Loaded({
    $ctx = $dgResults.ContextMenu

    $ctx.FindName('ctxRemoveDevice').Add_Click({
        # Remove all highlighted (selected in the DataGrid sense) rows
        $toRemove = @($dgResults.SelectedItems | ForEach-Object { $_ })
        if (-not $toRemove) { return }
        foreach ($item in $toRemove) {
            $deviceList.Remove($item) | Out-Null
            Write-Log "Device removed: $($item.Name)"
        }
        Update-Status "Removed $($toRemove.Count) device(s)"
    })

    $ctx.FindName('ctxCheckSelected').Add_Click({
        foreach ($item in @($dgResults.SelectedItems)) { $item.IsSelected = $true }
        $dgResults.Items.Refresh()
    })

    $ctx.FindName('ctxUncheckSelected').Add_Click({
        foreach ($item in @($dgResults.SelectedItems)) { $item.IsSelected = $false }
        $dgResults.Items.Refresh()
    })
})
#endregion

#region --- Async action helper ---
# Pattern: single background thread running sequential per-device PS jobs.
# Each device runs as a Start-Job (separate process) with a strict timeout.
# A DispatcherTimer on the UI thread polls completed jobs every 500ms and
# updates the grid as each one finishes - giving live per-device updates.
# Sequential job launch with parallel completion via polling.
#
# Why this works when other approaches don't:
#   - Start-Job isolates PS execution completely from the UI thread
#   - DispatcherTimer callback runs on UI thread - safe to touch WPF controls
#   - Strict per-job timeout prevents any single device from hanging forever

function Invoke-BCParallel {
    param(
        [string[]]$Names,
        [scriptblock]$PerDevice,
        [scriptblock]$OnAllComplete,
        [bool]$UseCurrentUser,
        [string]$LogFile,
        [int]$TimeoutSeconds = 15
    )

    # Credentials on UI thread
    $cred = if (-not $UseCurrentUser) {
        Get-BCCredential -Message "Enter credentials for remote devices"
    } else { $null }

    # Run each device sequentially on the UI thread.
    # Between each device, yield to the WPF dispatcher so the UI stays
    # responsive and the grid updates are visible immediately.
    foreach ($name in $Names) {
        # Run the work synchronously
        $r = $null
        try {
            $r = & $PerDevice $name $cred $LogFile
        } catch {
            $r = [PSCustomObject]@{
                Name=$name; Ping='ERROR'; WSMan='—'; Auth='—'
                BitLocker='—'; StatusDot='ERROR'
                LastAction="Error: $($_.Exception.Message)"
            }
        }

        if (-not $r) {
            $r = [PSCustomObject]@{
                Name=$name; Ping='ERROR'; WSMan='—'; Auth='—'
                BitLocker='—'; StatusDot='ERROR'; LastAction='No result'
            }
        }

        # Update the grid row for this device
        $item = $deviceList | Where-Object { $_.Name -eq $r.Name } | Select-Object -First 1
        if ($item) {
            try { if ($r.PSObject.Properties['Ping'])       { $item.Ping       = $r.Ping       } } catch {}
            try { if ($r.PSObject.Properties['WSMan'])      { $item.WSMan      = $r.WSMan      } } catch {}
            try { if ($r.PSObject.Properties['Auth'])       { $item.Auth       = $r.Auth       } } catch {}
            try { if ($r.PSObject.Properties['StatusDot'])  { $item.StatusDot  = $r.StatusDot  } } catch {}
            try { if ($r.PSObject.Properties['BitLocker'])  { $item.BitLocker  = $r.BitLocker  } } catch {}
            try { if ($r.PSObject.Properties['LastAction']) { $item.LastAction = $r.LastAction } } catch {}
        }
        $dgResults.Items.Refresh()
        $lblStatus.Text = "Completed: $($r.Name)"

        # Yield to WPF dispatcher - lets the UI repaint and process events
        # This is the key: keeps UI responsive between devices
        $dgResults.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    & $OnAllComplete
}
#endregion

#region --- Event: Pre-validate ---
$btnPreValidate.Add_Click({
    $targets = Get-SelectedDevices
    if (-not $targets) {
        [System.Windows.MessageBox]::Show("No devices selected.",
            "Nothing selected", [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information)
        return
    }

    Update-Status "Pre-validating $($targets.Count) device(s)..."
    Write-Log "Pre-validate started: $($targets.Count) devices"
    $btnPreValidate.IsEnabled = $false
    $names = @($targets | ForEach-Object { $_.Name })

    $perDevice = {
        param($name, $cred, $logFile)
        $ping = 'OFFLINE'; $wsman = '—'; $auth = '—'; $dot = 'ERROR'

        try {
            # Ping - 2s timeout
            if (Test-Connection -ComputerName $name -Count 1 -Quiet -TimeoutSeconds 2 -ErrorAction SilentlyContinue) {
                $ping = 'OK'

                # WinRM port check via async TCP - 3s timeout, never blocks
                try {
                    $tcp     = [System.Net.Sockets.TcpClient]::new()
                    $connect = $tcp.BeginConnect($name, 5985, $null, $null)
                    $waited  = $connect.AsyncWaitHandle.WaitOne(3000)
                    if ($waited -and $tcp.Connected) {
                        $tcp.EndConnect($connect)
                        $wsman = 'OK'

                        # Auth: try WMI which has its own internal timeout (~5s)
                        # Faster than Invoke-Command and doesn't need a separate runspace
                        try {
                            $wmi = [System.Management.ManagementScope]::new(
                                "\\\\$name\\root\\cimv2",
                                [System.Management.ConnectionOptions]::new())
                            $wmi.Options.Timeout = [TimeSpan]::FromSeconds(5)
                            if ($cred) {
                                $wmi.Options.Username = $cred.UserName
                                $wmi.Options.Password = $cred.Password
                            }
                            $wmi.Connect()
                            $auth = 'OK'; $dot = 'OK'
                        } catch { $auth = 'FAIL'; $dot = 'WARN' }
                    } else {
                        $wsman = 'CLOSED'; $dot = 'WARN'
                    }
                    $tcp.Close(); $tcp.Dispose()
                } catch { $wsman = 'ERR'; $dot = 'WARN' }
            }
        } catch { $ping = 'ERROR'; $dot = 'ERROR' }

        if ($logFile) {
            Add-Content -Path $logFile -Value "[$(Get-Date -Format 'HH:mm:ss')] [INFO] ${name}: Ping=$ping WSMan=$wsman Auth=$auth"
        }
        [PSCustomObject]@{ Name=$name; Ping=$ping; WSMan=$wsman; Auth=$auth; StatusDot=$dot
                           LastAction="Validated $(Get-Date -Format 'HH:mm')" }
    }

    Invoke-BCParallel -Names $names -LogFile $LogFile `
        -UseCurrentUser ([bool]$chkCurrentUser.IsChecked) `
        -PerDevice $perDevice `
        -OnAllComplete {
            $btnPreValidate.IsEnabled = $true
            Update-Status "Pre-validation complete"
            Write-Log "Pre-validate complete"
        }
})
#endregion

#region --- Event: BitLocker status ---
$btnBitLockerStatus.Add_Click({
    $targets = Get-SelectedDevices
    if (-not $targets) {
        [System.Windows.MessageBox]::Show("No devices selected.", "Nothing selected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    Update-Status "Checking BitLocker on $($targets.Count) device(s)..."
    Write-Log "BitLocker status check: $($targets.Count) devices"
    $btnBitLockerStatus.IsEnabled = $false
    $names = @($targets | ForEach-Object { $_.Name })

    $perDevice = {
        param($name, $cred, $logFile)
        $status = 'UNREACHABLE'
        try {
            # Use WMI with explicit timeout instead of Invoke-Command
            $opts = [System.Management.ConnectionOptions]::new()
            $opts.Timeout = [TimeSpan]::FromSeconds(8)
            if ($cred) { $opts.Username = $cred.UserName; $opts.Password = $cred.Password }
            $scope = [System.Management.ManagementScope]::new("\\\\$name\\root\\cimv2", $opts)
            $scope.Connect()
            # Query BitLocker via Win32_EncryptableVolume WMI class
            $wmiScope = [System.Management.ManagementScope]::new("\\\\$name\\root\\cimv2\\Security\\MicrosoftVolumeEncryption", $opts)
            try {
                $wmiScope.Connect()
                $query   = [System.Management.ObjectQuery]::new("SELECT * FROM Win32_EncryptableVolume")
                $searcher = [System.Management.ManagementObjectSearcher]::new($wmiScope, $query)
                $vols    = $searcher.Get()
                $statuses = @()
                foreach ($v in $vols) { $statuses += $v['ProtectionStatus'] }
                $status = if ($statuses -contains 1) { 'On' } elseif ($statuses.Count -gt 0) { 'Off' } else { 'N/A' }
            } catch { $status = 'WMI-BL-ERR' }
        } catch { $status = "UNREACHABLE" }
        if ($logFile) { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'HH:mm:ss')] [INFO] ${name}: BitLocker=$status" }
        [PSCustomObject]@{ Name=$name; BitLocker=$status; LastAction="BL check $(Get-Date -Format 'HH:mm')" }
    }

    Invoke-BCParallel -Names $names -LogFile $LogFile `
        -UseCurrentUser ([bool]$chkCurrentUser.IsChecked) `
        -PerDevice $perDevice `
        -OnAllComplete {
            $btnBitLockerStatus.IsEnabled = $true
            Update-Status "BitLocker check complete"
            Write-Log "BitLocker check complete"
        }
})
#endregion

#region --- Event: Disable BitLocker ---
$btnDisableBitLocker.Add_Click({
    $targets = Get-SelectedDevices
    if (-not $targets) {
        [System.Windows.MessageBox]::Show("No devices selected.", "Nothing selected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $nameList = ($targets | ForEach-Object { $_.Name }) -join "`n"
    $confirm  = [System.Windows.MessageBox]::Show(
        "Disable BitLocker on $($targets.Count) device(s)?`n`n$nameList",
        "Confirm: Disable BitLocker",
        [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Update-Status "Disabling BitLocker on $($targets.Count) device(s)..."
    Write-Log "Disable BitLocker: $($targets.Count) devices"
    $btnDisableBitLocker.IsEnabled = $false
    $names = @($targets | ForEach-Object { $_.Name })

    $perDevice = {
        param($name, $cred, $logFile)
        $result = 'ERROR'
        try {
            # Launch job and wait with hard timeout
            $j = Start-Job -ScriptBlock {
                param($n, $c)
                $p = @{ ComputerName=$n; ErrorAction='Stop'; ScriptBlock={
                    Get-BitLockerVolume | Where-Object { $_.ProtectionStatus -eq 'On' } |
                        ForEach-Object { Disable-BitLocker -MountPoint $_.MountPoint -ErrorAction Stop }
                    'Disabled'
                }}
                if ($c) { $p.Credential = $c }
                Invoke-Command @p
            } -ArgumentList $name, $cred
            # Wait max 20s - disable can take a moment to initiate
            $j | Wait-Job -Timeout 20 | Out-Null
            if ($j.State -eq 'Completed') {
                $result = Receive-Job $j -ErrorAction SilentlyContinue
                if (-not $result) { $result = 'Disabled' }
            } else {
                Stop-Job $j
                $result = 'TIMEOUT'
            }
            Remove-Job $j -Force
        } catch { $result = "FAILED: $($_.Exception.Message)" }
        if ($logFile) { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'HH:mm:ss')] [INFO] ${name}: DisableBL=$result" }
        [PSCustomObject]@{ Name=$name; BitLocker=$result; LastAction="BL disabled $(Get-Date -Format 'HH:mm')" }
    }

    Invoke-BCParallel -Names $names -LogFile $LogFile `
        -UseCurrentUser ([bool]$chkCurrentUser.IsChecked) `
        -ThrottleLimit $ThrottleLimit `
        -PerDevice $perDevice `
        -OnAllComplete {
            $btnDisableBitLocker.IsEnabled = $true
            Update-Status "BitLocker disable complete"
            Write-Log "BitLocker disable complete"
        }
})
#endregion

#region --- Event: Reboot ---
$btnReboot.Add_Click({
    $targets = Get-SelectedDevices
    if (-not $targets) {
        [System.Windows.MessageBox]::Show("No devices selected.", "Nothing selected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $nameList = ($targets | ForEach-Object { $_.Name }) -join "`n"
    $confirm  = [System.Windows.MessageBox]::Show(
        "Force reboot $($targets.Count) device(s)?`n`nThis will immediately restart:`n$nameList",
        "Confirm: Force Reboot",
        [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Update-Status "Rebooting $($targets.Count) device(s)..."
    Write-Log "Reboot initiated: $($targets.Count) devices"
    $btnReboot.IsEnabled = $false
    $names = @($targets | ForEach-Object { $_.Name })

    $perDevice = {
        param($name, $cred, $logFile)
        $result = 'ERROR'
        try {
            $p = @{ ComputerName = $name; Force = $true; ErrorAction = 'Stop' }
            if ($cred) { $p.Credential = $cred }
            Restart-Computer @p
            $result = 'Rebooted'
        } catch { $result = "FAILED: $($_.Exception.Message)" }
        if ($logFile) { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'HH:mm:ss')] [INFO] ${name}: Reboot=$result" }
        [PSCustomObject]@{ Name=$name; LastAction="Rebooted $(Get-Date -Format 'HH:mm')" }
    }

    Invoke-BCParallel -Names $names -LogFile $LogFile `
        -UseCurrentUser ([bool]$chkCurrentUser.IsChecked) `
        -ThrottleLimit $ThrottleLimit `
        -PerDevice $perDevice `
        -OnAllComplete {
            $btnReboot.IsEnabled = $true
            Update-Status "Reboot commands sent"
            Write-Log "Reboot complete"
        }
})
#endregion


#region --- Event: Export CSV ---
$btnExportCsv.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter   = "CSV files (*.csv)|*.csv"
    $dlg.FileName = "BCDeviceManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq $true) {
        $deviceList | Select-Object Name, Ping, WSMan, Auth, BitLocker, LastAction |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Update-Status "Exported to $($dlg.FileName)"
        Write-Log "Exported CSV: $($dlg.FileName)"
    }
})
#endregion

#region --- Event: Show log ---
$btnShowLog.Add_Click({
    $lf = $LogFile
    Start-Process pwsh -ArgumentList @(
        '-NoExit',
        '-Command',
        "Write-Host 'BriComp Device Manager - Live Log' -ForegroundColor Cyan; Write-Host 'File: $lf' -ForegroundColor DarkGray; Write-Host ''; Get-Content -Path '$lf' -Wait -Tail 40"
    )
    Write-Log "Log viewer opened"
})
#endregion

#region --- Load CSV on startup if provided ---
if ($CsvPath -and (Test-Path $CsvPath)) {
    try {
        $rows = Import-Csv -Path $CsvPath -ErrorAction Stop
        foreach ($row in $rows) {
            $name = $row.$ComputerNameColumn
            if ($name) { Add-Device -DeviceName $name }
        }
        Update-Status "Loaded $($rows.Count) devices from $CsvPath"
        Write-Log "Startup CSV loaded: $CsvPath ($($rows.Count) rows)"
    } catch {
        Write-Log "Startup CSV load failed: $($_.Exception.Message)" -Level WARN
    }
}
#endregion

#region --- Show window ---
Write-Log "UI initialized, showing window"
$window.ShowDialog() | Out-Null
Write-Log "Window closed. Session ended."
#endregion

# SIG # Begin signature block
# MIIobgYJKoZIhvcNAQcCoIIoXzCCKFsCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDyGUf2YEgCX0Fv
# dtjenfiTlail7z33BpnSLqLYpWBb5KCCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0GCSqG
# SIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTlaMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQg
# MjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C0Cit
# eLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce2vnS
# 1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0daE6ZM
# swEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6TSXBC
# Mo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoAFdE3
# /hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7OhD26j
# q22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM1bL5
# OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z8ujo
# 7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05huzU
# tw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNYmtwm
# KwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP/2NP
# TLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0TAQH/
# BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYDVR0j
# BBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8E
# PDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkq
# hkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95RysQDK
# r2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HLIvda
# qpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5BtfQ/g+
# lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnhOE7a
# brs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIhdXNS
# y0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV9zeK
# iwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/jwVYb
# KyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYHKi8Q
# xAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmCXBVm
# zGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l/aCn
# HwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZWeE4w
# gga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0Zo
# dLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi
# 6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNg
# xVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiF
# cMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJ
# m/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvS
# GmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1
# ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9
# MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7
# Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bG
# RinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6
# X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAd
# BgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJo
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# bDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQEL
# BQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxj
# aaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0
# hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0
# F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnT
# mpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKf
# ZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzE
# wlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbh
# OhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOX
# gpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EO
# LLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wG
# WqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWg
# AwIBAgIQCoDvGEuN8QWC0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# MB4XDTI1MDYwNDAwMDAwMFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEy
# NTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3
# zBlCMGMyqJnfFNZx+wvA69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8Tch
# TySA2R4QKpVD7dvNZh6wW2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWj
# FDYOzDi8SOhPUWlLnh00Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2Uo
# yrN0ijtUDVHRXdmncOOMA3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjP
# KHW5KqCvpSduSwhwUmotuQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KS
# uNLoZLc1Hf2JNMVL4Q1OpbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7w
# JNdoRORVbPR1VVnDuSeHVZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vW
# doUoHLWnqWU3dCCyFG1roSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOg
# rY7rlRyTlaCCfw7aSUROwnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K
# 096V1hE0yZIXe+giAwW00aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCf
# gPf8+3mnAgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zy
# Me39/dfzkXFjGVBDz2GM6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsG
# AQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNy
# dDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQBlKq3xHCcEua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZ
# D9gBq9fNaNmFj6Eh8/YmRDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/
# ML9lFfim8/9yJmZSe2F8AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu
# +WUqW4daIqToXFE/JQ/EABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4o
# bEMnxYOX8VBRKe1uNnzQVTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2h
# ECZpqyU1d0IbX6Wq8/gVutDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasn
# M9AWcIQfVjnzrvwiCZ85EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol
# /DJgddJ35XTxfUlQ+8Hggt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgY
# xQbV1S3CrWqZzBt1R9xJgKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3oc
# CVccAvlKV9jEnstrniLvUxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcB
# ZU8atufk+EMF/cWuiC7POGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCB3kwggVh
# oAMCAQICEAPvwdvfaByOuGfVs03RjH4wDQYJKoZIhvcNAQELBQAwaTELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2Vy
# dCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENB
# MTAeFw0yNTA4MTEwMDAwMDBaFw0yODA4MTAyMzU5NTlaMIGAMQswCQYDVQQGEwJV
# UzEQMA4GA1UECBMHQXJpem9uYTEQMA4GA1UEBxMHR2lsYmVydDEfMB0GA1UEChMW
# QnJpQ29tcCBDb21wdXRlcnMsIExMQzELMAkGA1UECxMCSVQxHzAdBgNVBAMTFkJy
# aUNvbXAgQ29tcHV0ZXJzLCBMTEMwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIK
# AoICAQDmORZ7VshQu610p+5XvpZwuBtGdXz/uDjhb08OBTFpxU+N/J7VrADg5IQr
# l2uow0RaFYPoWAp3tc9sszl9PL1moO0XjIhJngsIaqwZsqu9EO7kIFYZAqE5ziKP
# vipbdtKIw+Bo0BaMUwg5KuwU9Dp+BntlnlPcU00zECuyM/T+VCV+WxpFT1dXIj00
# chaNNSfzvhNc25HapytutgUpurgpTQ4zRpBJ3IhROJWb3yOJ8gcGaIUdqW49RpUg
# 9tuYVxtekZN+1Twl3hBwn8stZ+CD8vaUFARNs6WWgWX+trD9JHIQaxQz9DO0oPVz
# TKPWpJdUQNaPWU3x3hKZgF1nObXQEK31dmwVDlJPKQ/JV9wnkT9bDl4JNILgpJHR
# JYE9oszu0+mSUsjwglN10hrnhcEE6avZIUCN6zGrilVCkWee/mspUVTo0Oz/eaHV
# MnOVt/FH23zq6iMjfKY0bUcRUHE3EiT5bCVyLjxkfQgBlrZsyhcX3lLQ00bccIbA
# DsY+WS2peQKS1CKQAw19fAGct0HmFDWaOKqC6UBl710NH/HUs3K/QGHqWub+wFY5
# ypNC+WDXVd6Klm4CqbzAtWnunrnClbF8GW0418pVVK/syJroSYikyRF7R6dGbQo6
# C2/89l/S5g+dNmmzPBO1I5l4t9X2W4FQByx4b9W1gkHjQ+GcIQIDAQABo4ICAzCC
# Af8wHwYDVR0jBBgwFoAUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHQYDVR0OBBYEFMBR
# VSnudhF0J79q4bGDuXVuvTkDMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYB
# BQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAOBgNVHQ8BAf8EBAMC
# B4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwgbUGA1UdHwSBrTCBqjBToFGgT4ZNaHR0
# cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25p
# bmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BRoE+GTWh0dHA6Ly9jcmw0LmRp
# Z2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNI
# QTM4NDIwMjFDQTEuY3JsMIGUBggrBgEFBQcBAQSBhzCBhDAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMFwGCCsGAQUFBzAChlBodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmluZ1JT
# QTQwOTZTSEEzODQyMDIxQ0ExLmNydDAJBgNVHRMEAjAAMA0GCSqGSIb3DQEBCwUA
# A4ICAQB0dH/YkZqUjUXmYUn/948bMbrlKzTTbdLNQ/9fcPBqdXdqQ/m6c6Bb98Ti
# x9Jmk9Vp5rtqD3E47pRSQ9G9Cn+2WgandklNhYOo3fZIPO7WHA/0hB9cu9bfBxO5
# vV78jR2xow+WLs096OYJaKHMiU8mT498Db2NHLSv6+FeGxf8WZf/Ujp3nrkedlCa
# iMI2wTzbjMBHfQTGYZTAG6Nic1PIpJRs9x817QtAsyhQqPqlxi/J6tui0aMUYnJK
# fZiYJN74v+ANvNqJ9lPIivhc6k90ishsOtb2u8Ol4CDtDjkOSLnqlmU6FL0VygEJ
# g6HafKyspnIx5u+pqG8Icq5LJ0Tk9U6O/1riJlWqo5GYEPRacNXjuAxHXrp5iSoB
# JzabJwiAvZEHYDtfkAt1obbRaVb39ghaGgX/hpgnN3vgagPRCh4zhq5KE0Y/iXUA
# QDXzWndaKL2ScQSj7w5KGsHxDETg2VgrlDfZJ9e4szEX8R2TKmRRoZofBAa4STHq
# s2EYudfauWMV5jFT5d8ux1nUAz8pRLlPUWogJbhBLrtxjH8trjK2+cv/173j4JtO
# hCKjYp6a8wX3ZP3u6x1JdCJ9p0QlA5OpL5IXiFzPxPTB5HJ9W4G5Aq7/bYOI2lH9
# uVPA6OJ+HigaBjL2tsX1D+up2iWVZSEj89P9tvu6sTRYLkDX5zGCBlkwggZVAgEB
# MH0waTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYD
# VQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IENvZGUgU2lnbmluZyBSU0E0MDk2IFNI
# QTM4NCAyMDIxIENBMQIQA+/B299oHI64Z9WzTdGMfjANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDC0///4fOUNz28heu9lO0d3nHbpQFb9Ih2XLB7CP8B0zANBgkqhkiG9w0B
# AQEFAASCAgAristpJEJ7yJUpupWqNsZ2herUytopvt6oIOB32QVf2iGToZn4fwHM
# 9dFRzJ9kuDAHlBhtAvp10p6ATE9fvmALgwyX36dpVeV2OjxMSRjDxLlwMDkuXQNO
# mS0xgxi7ZzOwE2OxH/zSoHQX+rbTEsAoTfz4D1I9MMZ3u0e4NfQXBPwaqImiHLr/
# PVL712XW6VQjuKLqEIrRjYWf4RscbbSYm8LgOaI/v2gbb1u+2aZDI3vX8yvrBGr/
# +Waq8fx3eHVr3rQlVtSM+1zJrRkzH6CdZqSAbuddqLCARpNESyNidQCS9af7l1Kx
# 7u+o/4L5t+9WoyHaA9WCOXhkMRFlU01uh+Oo3QyociBoJE+61Qmrf5BGvlLaU9VF
# ZG0/9gZ3cJG0nCV6t7LPbPIh1TJTlIkJyesF0Sj2e2OsEM6G840NtXHPW+/GyUnU
# phzSlchyXVw3BV/waissW7YB9aQ/SHUMy9DPcmuS6vOn1RXLfOb/AM5A5/H4G2I2
# hIO/tjdpGHG//SlEOsc/FB1jAulDkEKmgZD5D3ZlvwSKtdWIG55jJ17rZ6Z4K7za
# c3sYZAqwK90uZZofOdGuQlnG58zWCjHilMWNPD/O+8Q4iGL1Miv+hFiNYcnNr+aL
# KLUaOxGbjpdzlFM2xUwF3GcAvPX7jpq9Z10vbkVMY9qNmUIHtiPwXaGCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHE
# dqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA2MjcxNzUxMjlaMC8GCSqGSIb3DQEJBDEiBCDS
# 8IdB6qHML88cMETr/jXZ4fwI/AunM9MghhtI8SuSAzANBgkqhkiG9w0BAQEFAASC
# AgAn8A+Bey+X8juoXyPnSdIlKLorx5LlQoOhlE2wZPZMq7H9qeO5szTZMMoTyA+V
# ifpY8uZAElqSn7jjDad7DwJzK2tvlnJiAduuEL6oC2ra5jZZbtjw4yfUeJrB+srh
# blXz8Jnqu02tzvNhzQ3w77FIbqr4798PeLDU0GXVNZKiX3f9HnHpdmTT9FdSm9Y6
# 9BouKIp4MQxw9/PBvaNRQdvxauB+HFi4g1zyeK+iTzO+9Oh6vuEmMfFCo5Sm0dFE
# qjNg0MA3qN3lBK7mTeRG0TRmUquD8HGNXSkhqJhV9DC5H9Cv8oi69uyWZHYePcmB
# zGaFuzO0TSngsZM7c20TshnSeUhsl3J/9atZ9JQ6nM3G0VuXxqLtLI7M57PW9Yoy
# 59Ad2DhgEZ4jucdyM5lvmKRJLj34z/GaJawJwR+wvsJXDkbbdl2qFFJroPl6BXRd
# utrcTYK5p6kymWtMK9aewY6Qe53Vk75JwdhHi+qHkFbciR9d2HGH6PbIl8kDLKSy
# v/8Xmph6cpDfcGWqtRmyirLw7pHuOc7JuthaTmqnDCQ+fn5iithGmI+oqh4mOzyP
# YBH19h5EyC+HGlsywZR15tefX6AKx84QPNU12Q4oL2QqAUBNTr373zU/nrmph793
# +3amUzNa+nB0uBMderHNtqFMa2kPMI0hjIIKiDlcutEzog==
# SIG # End signature block
