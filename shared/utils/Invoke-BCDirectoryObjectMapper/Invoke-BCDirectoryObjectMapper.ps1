#Requires -Version 7.0
<#
.SYNOPSIS
    Invoke-BCDirectoryObjectMapper.ps1
    WPF GUI for building Source/Target identity-mapping CSVs for migration
    waves, covering both devices and users, against Entra ID and/or
    on-premises Active Directory.

.DESCRIPTION
    Takes a list of identities - short computer names in Devices mode, or
    UserPrincipalNames in Users mode - via ad-hoc entry or CSV, and resolves
    each one against a chosen Source directory (Entra ID or Active Directory,
    always required) and an optional Target directory (None, Entra ID, or
    Active Directory - independently selectable from Source). Results export
    to a wave-mapping CSV in the standard migration object-mapping format:

        SourceObjectId,SourceUserPrincipalName,SourceSamAccountName,
        TargetObjectId,TargetUserPrincipalName,TargetSamAccountName,Comments

    Source and Target are looked up independently by matching the same
    identity name/UPN in each selected directory - no cross-referencing via
    synced-identity attributes (e.g. onPremisesSecurityIdentifier) is done.

    Output filename: <WaveName>_DEVICES_<yyyyMMdd_HHmmss>.csv or
    <WaveName>_USERS_<yyyyMMdd_HHmmss>.csv depending on the selected mode.

.PARAMETER CsvPath
    Optional path to a CSV file to pre-load on startup.

.PARAMETER IdentityColumn
    Column name in CSV holding the identity (device name or UPN). If blank
    or not found, the first column in the CSV is used automatically.

.PARAMETER LogRoot
    Directory for log output. Default: script directory.

.PARAMETER OutputRoot
    Default directory offered in the export Save dialog. Default: script directory.

.EXAMPLE
    pwsh -File .\Invoke-BCDirectoryObjectMapper.ps1
    pwsh -File .\Invoke-BCDirectoryObjectMapper.ps1 -CsvPath .\identities.csv

.NOTES
    BriComp Computers, LLC
    Version : 2.3.0
    Requires: PowerShell 7+
              Entra lookups: Microsoft.Graph.Authentication,
                Microsoft.Graph.Identity.DirectoryManagement (devices),
                Microsoft.Graph.Users (users) - auto-installed from PSGallery
                on request.
              AD lookups: ActiveDirectory module (RSAT) - NOT installable
                from PSGallery; the tool prompts with the correct
                Add-WindowsCapability / Install-WindowsFeature command if
                missing.
    Auth    : Entra - interactive delegated sign-in (Connect-MgGraph),
                requesting Device.Read.All and User.Read.All together so a
                single sign-in covers both modes without reconnecting.
              AD - uses the current logged-on user's Windows credentials
                (no separate sign-in); the AD Domain field targets the
                -Server parameter on Get-ADUser/Get-ADComputer.

    Facts verified against Microsoft Learn and current community
    documentation on 2026-07-07 before writing this version:
      - Get-ADUser / Get-ADComputer -Filter uses PowerShell expression
        syntax ("Property -eq 'value'"), NOT OData - different from Graph.
      - Both cmdlets' default output includes ObjectGUID, SamAccountName,
        and UserPrincipalName. AD computer SamAccountName always includes
        a trailing "$" (e.g. "COMPUTER01$") - this is normal AD behavior.
        This tool strips that trailing "$" before writing SamAccountName
        to the grid/CSV (user SamAccountName never has one, so this only
        affects Devices mode against Active Directory).
      - The -Server parameter (all AD cmdlets) accepts a domain FQDN or a
        specific DC hostname.
      - The ActiveDirectory module is a Windows PowerShell 5.1 binary
        module; on PS7 (Windows 10 1809+ / Server 2019+) Import-Module
        ActiveDirectory works automatically via built-in implicit-remoting
        compatibility (WinPSCompatSession) - a one-time warning is normal,
        not an error.
      - The ActiveDirectory module is NOT on PSGallery - it ships via RSAT
        (Add-WindowsCapability -Online -Name
        Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0 on client OS, or
        Install-WindowsFeature RSAT-AD-PowerShell on Windows Server). Both
        require local Administrator rights - this tool does not run
        elevated by default (only BCDeviceManager-style tools that need
        remote actions on every launch do that), so RSAT install is
        offered on-demand via a one-time UAC-elevated child PowerShell
        process (Start-Process -Verb RunAs), not silently/automatically.
      - Win32_OperatingSystem.ProductType (via Get-CimInstance, built into
        PowerShell's CimCmdlets module) reliably distinguishes Windows
        client (1) from Server/DC (2 or 3) for picking the correct RSAT
        install command.
      - System.Windows.Clipboard.SetText works from this script because
        WPF (XamlReader.Load/ShowDialog) requires an STA thread, which
        this script's successful use of Window/MessageBox already
        confirms it's running under.
      - Get-MgUser ships in Microsoft.Graph.Users. Get-MgUser -UserId <UPN>
        does a direct identity fetch (single result or a terminating 404),
        avoiding the ambiguity a -Filter search could return.
      - OnPremisesSamAccountName is a real, retrievable Graph user property
        for hybrid-synced users, but is deliberately NOT used (see 2.2.1
        below) - SamAccountName is only ever populated from an actual AD
        lookup.
      - $env:USERDNSDOMAIN holds the current logged-on user's DNS domain -
        used to pre-fill the AD Domain field.

    === Revision History ===
    1.0.0  (2026-07-06) Initial release - Entra-only device lookup.
    2.0.0  (2026-07-07) Added Devices/Users mode toggle; independent
                         Source/Target directory selection (Entra ID or
                         Active Directory, Target optional); AD domain
                         field with auto-detect; AD lookups via
                         Get-ADUser/Get-ADComputer; CSV filename now
                         <Wave>_DEVICES_<timestamp>.csv or
                         <Wave>_USERS_<timestamp>.csv.
    2.0.1  (2026-07-07) Strip trailing "$" from AD computer SamAccountName
                         before writing it to the grid/CSV.
    2.0.2  (2026-07-07) Guard Add-Type with a type-existence check so
                         re-running the script in the same PowerShell
                         session/console doesn't throw "Cannot add type.
                         The type name 'BriComp.WaveItem' already exists."
    2.0.3  (2026-07-07) Fixed Connections status (Entra ID / Active
                         Directory) being scrolled out of view by default
                         - it's now pinned outside the scrollable sidebar
                         area. Each line is now also only shown when that
                         directory is currently selected as Source and/or
                         Target.
    2.0.4  (2026-07-07) Replaced the 2.0.2 Add-Type "already exists" guard
                         (which could silently leave a stale, older-shaped
                         type active and cause cryptic property-not-found
                         errors, e.g. StatusColor) with a per-run-unique
                         type name (GUID-suffixed), which eliminates the
                         Add-Type collision entirely rather than papering
                         over it.
    2.1.0  (2026-07-07) When the ActiveDirectory module is missing, the
                         tool now detects Windows client vs. Server (via
                         Win32_OperatingSystem.ProductType) and shows only
                         the applicable RSAT install command, with a Copy
                         button and an "Install Now (Admin)" button that
                         launches a one-time UAC-elevated PowerShell to run
                         it - full silent auto-install was deliberately not
                         done since it would require this tool to always
                         run elevated (UAC prompt on every launch, even for
                         users who never touch AD lookups).
    2.1.1  (2026-07-07) Renamed from Invoke-BCEntraObjectMapper.ps1 to
                         Invoke-BCDirectoryObjectMapper.ps1 - the old name
                         no longer reflected the tool's scope once
                         Active Directory support was added in 2.0.0.
                         No functional changes.
    2.2.0  (2026-07-07) Toggling between Devices and Users now prompts for
                         confirmation before clearing the current list (no
                         prompt if the list is already empty). Cancel
                         reverts the radio button selection and leaves the
                         list untouched. Implemented via the RadioButton
                         Click event rather than Checked, since Click is
                         only raised by real user interaction - not by the
                         programmatic IsChecked reassignment used to revert
                         a canceled switch - avoiding any risk of recursion.
    2.2.1  (2026-07-07) Export-Csv now uses -UseQuotes AsNeeded instead of
                         the PS7 default of -UseQuotes Always, so fields
                         are only quoted if they actually contain a comma,
                         quote character, or newline. Entra user lookups no
                         longer populate SamAccountName from
                         OnPremisesSamAccountName - SamAccountName in the
                         grid/CSV is now only ever populated by an actual
                         Active Directory lookup, regardless of hybrid sync
                         status.
    2.3.0  (2026-07-07) CSV header/column set now differs by mode, per
                         exact spec: Users exports all 7 columns with
                         "(Optional)" suffixes on every field except
                         SourceObjectId; Devices exports only 3 columns
                         (SourceObjectId, "SourceSamAccountName
                         (Optional)", "Comments (Optional)") - Source UPN
                         and all Target fields are dropped from the
                         Devices CSV even if a Target directory was
                         resolved (that data still shows in the grid, just
                         isn't exported). Header text is produced by
                         naming the PSCustomObject properties exactly as
                         specified, rather than via literal quote
                         characters - -UseQuotes AsNeeded (2.2.1) still
                         applies to both, so nothing is quoted in the
                         output unless a value actually needs it.
#>

[CmdletBinding()]
param(
    [string]$CsvPath        = '',
    [string]$IdentityColumn = '',
    [string]$LogRoot        = '',
    [string]$OutputRoot     = ''
)

#region --- Assemblies ---
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Runtime
Add-Type -AssemblyName System.DirectoryServices.AccountManagement
[System.Collections.ObjectModel.ObservableCollection[object]] | Out-Null
#endregion

#region --- Logging setup ---
if (-not $LogRoot) { $LogRoot = $PSScriptRoot }
if (-not $OutputRoot) { $OutputRoot = $PSScriptRoot }
$LogDir  = Join-Path $LogRoot ("Logs_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$LogFile = Join-Path $LogDir 'BCDirectoryObjectMapper.log'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

Write-Log "BriComp Directory Object Mapper v2.1.1 started"
Write-Log "Log directory: $LogDir"
$script:GraphConnected = $false
$script:ADAvailable    = $false
$script:CurrentMode    = 'Devices'
#endregion

#region --- XAML UI definition ---
[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BriComp Identity Object Mapper"
    Height="680" Width="1180"
    MinHeight="560" MinWidth="900"
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
    <Style TargetType="ComboBox">
      <Setter Property="Padding" Value="5,4"/>
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
      <ToolBar Band="1" BandIndex="1" Background="#F0F0F0">

        <Button Name="btnConnect" ToolTip="Connect to the directories selected below (Entra ID sign-in and/or Active Directory check)">
          <TextBlock Text="Connect" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnLookup" ToolTip="Resolve Source/Target identities for selected rows">
          <TextBlock Text="Lookup Identities" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnExportCsv" ToolTip="Export the wave mapping CSV">
          <TextBlock Text="Export Wave CSV" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnShowLog" ToolTip="Open live log viewer">
          <TextBlock Text="Show log" VerticalAlignment="Center"/>
        </Button>

      </ToolBar>

      <ToolBar Band="2" BandIndex="1" Background="#F0F0F0">

        <TextBlock Text="Select:" VerticalAlignment="Center" Foreground="#888888"
                   FontSize="11" Margin="4,0,6,0"/>

        <Button Name="btnCheckAll" ToolTip="Check all rows">
          <TextBlock Text="All" VerticalAlignment="Center"/>
        </Button>

        <Button Name="btnUncheckAll" ToolTip="Uncheck all rows">
          <TextBlock Text="None" VerticalAlignment="Center"/>
        </Button>

        <Separator/>

        <Button Name="btnUncheckNotFound" ToolTip="Uncheck rows that were not found or are ambiguous (run Lookup first)">
          <TextBlock Text="Uncheck not found" VerticalAlignment="Center"/>
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
        <TextBlock Name="lblCounts" Text="0 rows" FontSize="11" Foreground="#555555"/>
      </StatusBarItem>
      <Separator/>
      <StatusBarItem>
        <TextBlock Name="lblLogPath" Text="" FontSize="11" Foreground="#888888"/>
      </StatusBarItem>
    </StatusBar>

    <!-- Main body: sidebar + grid -->
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="270" MinWidth="220"/>
        <ColumnDefinition Width="4"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <!-- Sidebar -->
      <DockPanel Grid.Column="0" Background="#F8F8F8">

        <ScrollViewer DockPanel.Dock="Top" VerticalScrollBarVisibility="Auto" MaxHeight="420">
        <Border BorderBrush="#E0E0E0" BorderThickness="0,0,0,1" Padding="8">
          <StackPanel>

            <!-- Wave name (required) -->
            <TextBlock Foreground="#888888" FontSize="10" FontWeight="SemiBold" Margin="0,0,0,2">
              <Run Text="WAVE NAME"/>
              <Run Text=" *" Foreground="#B91C1C"/>
            </TextBlock>
            <TextBlock Text="Required. Used to name the exported CSV." FontSize="10" Foreground="#AAAAAA" Margin="0,0,0,4" TextWrapping="Wrap"/>
            <Grid Margin="0,0,0,10">
              <TextBox Name="txtWaveName" ToolTip="e.g. Wave1-CSG-Migration"/>
              <TextBlock Name="txtWaveNamePH" Text="e.g. Wave1-CSG"
                         IsHitTestVisible="False" VerticalAlignment="Center"
                         Margin="6,0,0,0" Foreground="#AAAAAA" FontStyle="Italic" FontSize="11"/>
            </Grid>

            <!-- Mode -->
            <TextBlock Text="TYPE" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,4"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
              <RadioButton Name="rbModeDevices" GroupName="Mode" Content="Devices" IsChecked="True" Margin="0,0,16,0" VerticalAlignment="Center"/>
              <RadioButton Name="rbModeUsers" GroupName="Mode" Content="Users" VerticalAlignment="Center"/>
            </StackPanel>

            <!-- Directories -->
            <TextBlock Text="DIRECTORIES" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,4"/>
            <TextBlock Foreground="#888888" FontSize="10" Margin="0,0,0,2">
              <Run Text="Source"/>
              <Run Text=" *" Foreground="#B91C1C"/>
            </TextBlock>
            <ComboBox Name="cmbSourceDirectory" Margin="0,0,0,8" SelectedIndex="0">
              <ComboBoxItem Content="Entra ID" Tag="Entra"/>
              <ComboBoxItem Content="Active Directory" Tag="AD"/>
            </ComboBox>
            <TextBlock Text="Target (optional)" FontSize="10" Foreground="#888888" Margin="0,0,0,2"/>
            <ComboBox Name="cmbTargetDirectory" Margin="0,0,0,10" SelectedIndex="0">
              <ComboBoxItem Content="(None)" Tag="None"/>
              <ComboBoxItem Content="Entra ID" Tag="Entra"/>
              <ComboBoxItem Content="Active Directory" Tag="AD"/>
            </ComboBox>

            <!-- AD Domain -->
            <TextBlock Text="AD DOMAIN" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,2"/>
            <TextBlock Text="Used for any Active Directory lookups (Source and/or Target). Auto-filled from your current logon domain." FontSize="10" Foreground="#AAAAAA" Margin="0,0,0,4" TextWrapping="Wrap"/>
            <Grid Margin="0,0,0,10">
              <TextBox Name="txtADDomain" ToolTip="e.g. bricomp.com"/>
              <TextBlock Name="txtADDomainPH" Text="e.g. bricomp.com"
                         IsHitTestVisible="False" VerticalAlignment="Center"
                         Margin="6,0,0,0" Foreground="#AAAAAA" FontStyle="Italic" FontSize="11"/>
            </Grid>

            <TextBlock Text="TARGETS" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,6"/>

            <!-- Manual entry with watermark placeholder -->
            <DockPanel Margin="0,0,0,4">
              <Button DockPanel.Dock="Right" Name="btnAddDevice" Width="36" Padding="0" Margin="4,0,0,0" ToolTip="Add identity" FontSize="16" FontWeight="Bold">+</Button>
              <Grid>
                <TextBox Name="txtIdentity" AcceptsReturn="False" ToolTip="Enter a short computer name and press Enter or click +"/>
                <TextBlock Name="txtIdentityPH" Text="Enter device name..."
                           IsHitTestVisible="False" VerticalAlignment="Center"
                           Margin="6,0,0,0" Foreground="#AAAAAA" FontStyle="Italic" FontSize="11"/>
              </Grid>
            </DockPanel>

            <!-- CSV load + clear -->
            <DockPanel Margin="0,0,0,0">
              <Button DockPanel.Dock="Right" Name="btnClearAll" Width="60" Padding="4,0" Margin="4,0,0,0"
                      ToolTip="Clear all rows" Foreground="#B91C1C" FontSize="11">Clear</Button>
              <Button Name="btnLoadCsv" HorizontalAlignment="Stretch" FontSize="11">Load CSV...</Button>
            </DockPanel>

          </StackPanel>
        </Border>
        </ScrollViewer>

        <!-- Connection status - pinned, always visible (not inside the scroll area above),
             so it's never scrolled out of view. Each line's visibility is toggled at
             runtime based on whether Source and/or Target is currently set to that
             directory. -->
        <Border DockPanel.Dock="Top" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1" Padding="8" Background="#F0F0F0">
          <StackPanel>
            <TextBlock Text="CONNECTIONS" FontSize="10" FontWeight="SemiBold" Foreground="#888888" Margin="0,0,0,4"/>
            <TextBlock Name="lblGraphStatus" Text="Entra ID: Not connected" FontSize="11" Foreground="#B91C1C" TextWrapping="Wrap" Margin="0,0,0,4" Visibility="Collapsed"/>
            <TextBlock Name="lblADStatus" Text="Active Directory: Not connected" FontSize="11" Foreground="#B91C1C" TextWrapping="Wrap" Visibility="Collapsed"/>
            <TextBlock Name="lblNoConnectionsNeeded" Text="(Select Entra ID or Active Directory as Source/Target above)" FontSize="10" Foreground="#AAAAAA" TextWrapping="Wrap"/>
          </StackPanel>
        </Border>

        <!-- Identity list -->
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
                        <DataTrigger Binding="{Binding StatusColor}" Value="OK">
                          <Setter Property="Fill" Value="#16A34A"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding StatusColor}" Value="WARN">
                          <Setter Property="Fill" Value="#D97706"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding StatusColor}" Value="ERROR">
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
      <DataGrid Name="dgResults" Grid.Column="2" Margin="0" HorizontalScrollBarVisibility="Auto">

        <DataGrid.ContextMenu>
          <ContextMenu>
            <MenuItem Name="ctxRemoveDevice" Header="Remove row" />
            <Separator/>
            <MenuItem Name="ctxCheckSelected"   Header="Check selected rows" />
            <MenuItem Name="ctxUncheckSelected" Header="Uncheck selected rows" />
          </ContextMenu>
        </DataGrid.ContextMenu>

        <DataGrid.Columns>
          <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged}" Width="36" Header=""/>
          <DataGridTextColumn Binding="{Binding Name}"                     Header="Identity"                Width="150" IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding LookupStatus}"             Header="Status"                  Width="220" IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding SourceObjectId}"           Header="SourceObjectId"          Width="220" IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding SourceUserPrincipalName}"  Header="Source UPN"              Width="170" IsReadOnly="False"/>
          <DataGridTextColumn Binding="{Binding SourceSamAccountName}"     Header="Source SAM"              Width="120" IsReadOnly="False"/>
          <DataGridTextColumn Binding="{Binding TargetObjectId}"           Header="TargetObjectId"          Width="220" IsReadOnly="True"/>
          <DataGridTextColumn Binding="{Binding TargetUserPrincipalName}"  Header="Target UPN"              Width="170" IsReadOnly="False"/>
          <DataGridTextColumn Binding="{Binding TargetSamAccountName}"     Header="Target SAM"              Width="120" IsReadOnly="False"/>
          <DataGridTextColumn Binding="{Binding Comments}"                 Header="Comments"                Width="*"   IsReadOnly="False"/>
        </DataGrid.Columns>

      </DataGrid>

    </Grid>
  </DockPanel>
</Window>
'@
#endregion

#region --- Data model ---
# PowerShell classes cannot implement INotifyPropertyChanged reliably.
# Define the class in inline C# instead so WPF data binding works correctly.
#
# NOTE ON A REAL BUG FOUND IN 2.0.2: Add-Type permanently loads a compiled
# type into the current PowerShell process; .NET does not allow redefining
# a type of the same name within that process afterward. 2.0.2 "fixed" the
# resulting "Cannot add type. The type name already exists" error by
# skipping Add-Type when the name was already loaded - but that silently
# left an OLDER, differently-shaped version of the type active if an
# earlier script version (or an earlier run before an edit) had already
# loaded it in that same console session. That is exactly what produced
# "The property 'StatusColor' cannot be found on this object": the skip
# fired, the stale pre-2.0 shape stayed active, and StatusColor didn't
# exist on it. Verified as a known, documented Add-Type limitation with no
# way to unload/redefine a type in-process short of a new process.
#
# THE FIX: generate a fresh, process-and-run-unique class name every time
# this script executes (via a GUID), so Add-Type can never collide with a
# previously loaded type from an earlier run in the same session, and the
# skip-and-go-stale failure mode above becomes structurally impossible -
# not just detected, but eliminated. The WPF grid/list bind to properties
# by name via reflection at runtime regardless of the collection's
# declared generic type, so using ObservableCollection[object] instead of
# a compile-time-fixed ObservableCollection[BriComp.WaveItem] works fine
# and is required here since the class name is no longer known until now.
$script:WaveItemTypeName = "WaveItem_" + ([Guid]::NewGuid().ToString('N'))

try {
    Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;

namespace BriComp {
    public class $($script:WaveItemTypeName) : INotifyPropertyChanged {
        public event PropertyChangedEventHandler PropertyChanged;

        private void Notify(string prop) {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private string _name                     = "";
        private bool   _isSelected               = true;
        private string _lookupStatus             = "Not looked up";
        private string _statusColor              = "NONE";
        private string _sourceObjectId           = "";
        private string _sourceUserPrincipalName  = "";
        private string _sourceSamAccountName     = "";
        private string _targetObjectId           = "";
        private string _targetUserPrincipalName  = "";
        private string _targetSamAccountName     = "";
        private string _comments                 = "";

        public string Name                    { get { return _name;                    } set { _name                    = value; Notify("Name");                    } }
        public bool   IsSelected              { get { return _isSelected;              } set { _isSelected              = value; Notify("IsSelected");              } }
        public string LookupStatus            { get { return _lookupStatus;            } set { _lookupStatus            = value; Notify("LookupStatus");            } }
        public string StatusColor             { get { return _statusColor;             } set { _statusColor             = value; Notify("StatusColor");             } }
        public string SourceObjectId          { get { return _sourceObjectId;          } set { _sourceObjectId          = value; Notify("SourceObjectId");          } }
        public string SourceUserPrincipalName { get { return _sourceUserPrincipalName; } set { _sourceUserPrincipalName = value; Notify("SourceUserPrincipalName"); } }
        public string SourceSamAccountName    { get { return _sourceSamAccountName;    } set { _sourceSamAccountName    = value; Notify("SourceSamAccountName");    } }
        public string TargetObjectId          { get { return _targetObjectId;          } set { _targetObjectId          = value; Notify("TargetObjectId");          } }
        public string TargetUserPrincipalName { get { return _targetUserPrincipalName; } set { _targetUserPrincipalName = value; Notify("TargetUserPrincipalName"); } }
        public string TargetSamAccountName    { get { return _targetSamAccountName;    } set { _targetSamAccountName    = value; Notify("TargetSamAccountName");    } }
        public string Comments                { get { return _comments;                } set { _comments                = value; Notify("Comments");                } }
    }
}
"@ -ErrorAction Stop
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to compile the internal data type:`n$($_.Exception.Message)",
        "Startup error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    exit 1
}
$script:WaveItemFullName = "BriComp.$($script:WaveItemTypeName)"
#endregion

#region --- Build window ---
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$btnConnect        = $window.FindName('btnConnect')
$btnLookup         = $window.FindName('btnLookup')
$btnExportCsv      = $window.FindName('btnExportCsv')
$btnShowLog        = $window.FindName('btnShowLog')
$btnCheckAll       = $window.FindName('btnCheckAll')
$btnUncheckAll     = $window.FindName('btnUncheckAll')
$btnUncheckNotFound= $window.FindName('btnUncheckNotFound')
$btnAddDevice      = $window.FindName('btnAddDevice')
$txtIdentityPH     = $window.FindName('txtIdentityPH')
$txtWaveNamePH     = $window.FindName('txtWaveNamePH')
$txtADDomainPH     = $window.FindName('txtADDomainPH')
$btnClearAll       = $window.FindName('btnClearAll')
$btnLoadCsv        = $window.FindName('btnLoadCsv')
$txtIdentity       = $window.FindName('txtIdentity')
$txtWaveName       = $window.FindName('txtWaveName')
$txtADDomain       = $window.FindName('txtADDomain')
$rbModeDevices     = $window.FindName('rbModeDevices')
$rbModeUsers       = $window.FindName('rbModeUsers')
$cmbSourceDirectory= $window.FindName('cmbSourceDirectory')
$cmbTargetDirectory= $window.FindName('cmbTargetDirectory')
$lstDevices        = $window.FindName('lstDevices')
$dgResults         = $window.FindName('dgResults')
$lblStatus         = $window.FindName('lblStatus')
$lblCounts         = $window.FindName('lblCounts')
$lblLogPath        = $window.FindName('lblLogPath')
$lblGraphStatus    = $window.FindName('lblGraphStatus')
$lblADStatus       = $window.FindName('lblADStatus')
$lblNoConnectionsNeeded = $window.FindName('lblNoConnectionsNeeded')

$lblLogPath.Text = "Log: $LogFile"

$identityList = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$dgResults.ItemsSource  = $identityList
$lstDevices.ItemsSource = $identityList

# Auto-detect the current logon domain for the AD Domain field
$detectedDomain = $env:USERDNSDOMAIN
if (-not $detectedDomain) {
    try {
        $detectedDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name
    } catch {
        $detectedDomain = ''
    }
}
if ($detectedDomain) { $txtADDomain.Text = $detectedDomain }
#endregion

#region --- Helper functions ---
function Update-Status {
    param([string]$Message)
    $window.Dispatcher.Invoke([Action]{
        $lblStatus.Text = $Message
        $ok    = @($identityList | Where-Object { $_.StatusColor -eq 'OK' }).Count
        $other = @($identityList | Where-Object { $_.StatusColor -eq 'WARN' -or $_.StatusColor -eq 'ERROR' }).Count
        $total = $identityList.Count
        $lblCounts.Text = "$total rows - $ok resolved - $other not resolved"
    })
}

function Get-CurrentMode {
    if ($rbModeUsers.IsChecked) { return 'Users' }
    return 'Devices'
}

function Get-DirectoryTag {
    param($ComboBox)
    $item = $ComboBox.SelectedItem
    if ($item -and $item.Tag) { return $item.Tag.ToString() }
    return 'None'
}

function Update-ConnectionVisibility {
    # Shows the Entra ID / Active Directory connection status lines only
    # when that directory is currently selected as Source and/or Target,
    # so the sidebar always reflects what's actually relevant right now.
    $srcDir = Get-DirectoryTag -ComboBox $cmbSourceDirectory
    $tgtDir = Get-DirectoryTag -ComboBox $cmbTargetDirectory
    $needEntra = ($srcDir -eq 'Entra' -or $tgtDir -eq 'Entra')
    $needAD    = ($srcDir -eq 'AD'    -or $tgtDir -eq 'AD')

    $lblGraphStatus.Visibility = if ($needEntra) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $lblADStatus.Visibility    = if ($needAD)    { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $lblNoConnectionsNeeded.Visibility = if ($needEntra -or $needAD) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
}

function Add-Identity {
    param([string]$IdentityName)
    $name = $IdentityName.Trim()
    if (-not $name) { return }

    if ($identityList | Where-Object { $_.Name -eq $name }) {
        Update-Status "Already in list: $name"
        return
    }

    $item = New-Object -TypeName $script:WaveItemFullName
    $item.Name = $name
    $window.Dispatcher.Invoke([Action]{ $identityList.Add($item) })
    Write-Log "Identity added: $name"
    Update-Status "Added: $name"
}

function Get-SelectedIdentities {
    return @($identityList | Where-Object { $_.IsSelected })
}

function Test-BCGraphModules {
    param([string[]]$Required)
    $missing = @()
    foreach ($m in $Required) {
        if (-not (Get-Module -ListAvailable -Name $m)) { $missing += $m }
    }
    return $missing
}
#endregion

#region --- Helper: RSAT install dialog ---
function Get-BCWindowsOsType {
    # Win32_OperatingSystem.ProductType (documented WMI/CIM property):
    #   1 = Workstation (client, e.g. Windows 10/11)
    #   2 = Domain Controller
    #   3 = Server (non-DC)
    # Domain Controllers already ship the ActiveDirectory module, but a DC
    # would still use the server-style Install-WindowsFeature command if
    # it were somehow missing, so 2 and 3 are both treated as 'Server'.
    try {
        $productType = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).ProductType
        if ($productType -eq 1) { return 'Client' }
        return 'Server'
    } catch {
        Write-Log "OS type detection failed: $($_.Exception.Message)" -Level WARN
        return 'Unknown'
    }
}

function Get-BCRsatInstallCommand {
    param([string]$OsType)
    switch ($OsType) {
        'Client' { return 'Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' }
        'Server' { return 'Install-WindowsFeature RSAT-AD-PowerShell' }
        default  { return 'Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' }
    }
}

function Show-BCRsatInstallDialog {
    param([System.Windows.Window]$Owner)

    $osType = Get-BCWindowsOsType
    $cmd    = Get-BCRsatInstallCommand -OsType $osType

    [xml]$dlgXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Active Directory module missing"
        Width="540" SizeToContent="Height"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="12">
  <StackPanel Margin="18">
    <TextBlock Text="The Active Directory PowerShell module (RSAT) is not installed on this machine." TextWrapping="Wrap" Margin="0,0,0,10"/>
    <TextBlock Name="lblOsDetected" Text="" FontSize="11" Foreground="#888888" TextWrapping="Wrap" Margin="0,0,0,10"/>
    <TextBlock Text="This requires local Administrator rights. Either copy the command below into an elevated PowerShell window yourself, or click Install Now to launch one:" TextWrapping="Wrap" Margin="0,0,0,6"/>
    <TextBox Name="txtCommand" IsReadOnly="True" TextWrapping="Wrap" Padding="8" Background="#F5F5F5" BorderBrush="#CCCCCC" BorderThickness="1" Margin="0,0,0,6" FontFamily="Consolas"/>
    <TextBlock Name="lblCopied" Text=" " FontSize="11" Foreground="#16A34A" TextWrapping="Wrap" Margin="0,0,0,10"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btnInstallNow" Content="Install Now (Admin)" Width="150" Margin="0,0,8,0" Padding="8,5"/>
      <Button Name="btnCopy" Content="Copy" Width="80" Margin="0,0,8,0" Padding="8,5"/>
      <Button Name="btnClose" Content="Close" Width="80" IsDefault="True" IsCancel="True" Padding="8,5"/>
    </StackPanel>
  </StackPanel>
</Window>
'@

    $dlgReader = [System.Xml.XmlNodeReader]::new($dlgXaml)
    $dlgWin    = [System.Windows.Markup.XamlReader]::Load($dlgReader)
    if ($Owner) { $dlgWin.Owner = $Owner }

    $lblOsDetected = $dlgWin.FindName('lblOsDetected')
    $txtCommand    = $dlgWin.FindName('txtCommand')
    $lblCopied     = $dlgWin.FindName('lblCopied')
    $btnInstallNow = $dlgWin.FindName('btnInstallNow')
    $btnCopy       = $dlgWin.FindName('btnCopy')
    $btnClose      = $dlgWin.FindName('btnClose')

    $lblOsDetected.Text = switch ($osType) {
        'Client' { 'Detected: Windows client (10/11) - showing the client install command.' }
        'Server' { 'Detected: Windows Server - showing the server install command.' }
        default  { 'Could not detect Windows client vs. Server (showing the client command as a default).' }
    }
    $txtCommand.Text = $cmd

    $btnCopy.Add_Click({
        try {
            [System.Windows.Clipboard]::SetText($txtCommand.Text)
            $lblCopied.Text = "Copied to clipboard."
            $lblCopied.Foreground = [System.Windows.Media.Brushes]::DarkGreen
        } catch {
            $lblCopied.Text = "Could not copy to clipboard: $($_.Exception.Message)"
            $lblCopied.Foreground = [System.Windows.Media.Brushes]::Firebrick
        }
    })

    $btnInstallNow.Add_Click({
        try {
            $installCmd = $txtCommand.Text
            $inner = "Write-Host 'Installing Active Directory PowerShell module...' -ForegroundColor Cyan; $installCmd; Write-Host ''; Write-Host 'Done - close this window, then restart the mapper tool.' -ForegroundColor Green"
            Start-Process -FilePath pwsh -Verb RunAs -ArgumentList @('-NoExit', '-Command', $inner) -ErrorAction Stop
            $lblCopied.Text = "Elevated PowerShell launched - approve the UAC prompt, then restart this tool once it finishes."
            $lblCopied.Foreground = [System.Windows.Media.Brushes]::DarkGreen
        } catch {
            $lblCopied.Text = "Could not launch elevated install: $($_.Exception.Message)"
            $lblCopied.Foreground = [System.Windows.Media.Brushes]::Firebrick
        }
    })

    $btnClose.Add_Click({ $dlgWin.Close() })

    $dlgWin.ShowDialog() | Out-Null
}
#endregion

#region --- Event: Mode toggle (confirm before clearing the list) ---
# NOTE: Using Click (not Checked) deliberately. WPF's ToggleButton/RadioButton
# raises Checked/Unchecked whenever IsChecked changes for ANY reason,
# including a programmatic assignment - but Click is raised only by real
# user interaction (ButtonBase.OnClick), never by setting IsChecked in code
# (confirmed against Microsoft Learn's ToggleButton reference and third-
# party WPF control docs, which consistently describe Click as a distinct,
# interaction-only event separate from the state-driven Checked/Unchecked/
# Activate events). That means reverting the selection below (when the
# person cancels) via $rbModeX.IsChecked = $true is safe and cannot
# recursively re-trigger this same Click handler.

function Set-BCModeUI {
    param([string]$Mode)
    if ($Mode -eq 'Devices') {
        $txtIdentityPH.Text = 'Enter device name...'
        $txtIdentity.ToolTip = 'Enter a short computer name and press Enter or click +'
        Update-Status 'Mode: Devices'
    } else {
        $txtIdentityPH.Text = 'Enter user UPN (user@domain.com)...'
        $txtIdentity.ToolTip = 'Enter the user UserPrincipalName (UPN) and press Enter or click +'
        Update-Status 'Mode: Users'
    }
    $script:CurrentMode = $Mode
}

function Request-BCModeSwitch {
    param([string]$NewMode, [System.Windows.Controls.RadioButton]$RevertRadioButton)

    if ($script:CurrentMode -eq $NewMode) { return }

    if ($identityList.Count -eq 0) {
        Set-BCModeUI -Mode $NewMode
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Switching between Devices and Users clears the current list.`n`n$($identityList.Count) row(s) will be removed. Continue?",
        "Switch mode - list will be cleared",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning)

    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        $removed = $identityList.Count
        $identityList.Clear()
        Set-BCModeUI -Mode $NewMode
        Write-Log "Mode switched to $NewMode - cleared $removed row(s)"
    } else {
        # Revert the visual selection back to the mode that's still active.
        # This only raises Checked/Unchecked (see note above), not Click,
        # so it will not re-enter this handler.
        $RevertRadioButton.IsChecked = $true
        Update-Status "Mode switch canceled - staying on $($script:CurrentMode)"
    }
}

$rbModeDevices.Add_Click({
    Request-BCModeSwitch -NewMode 'Devices' -RevertRadioButton $rbModeUsers
})
$rbModeUsers.Add_Click({
    Request-BCModeSwitch -NewMode 'Users' -RevertRadioButton $rbModeDevices
})
#endregion

#region --- Event: Directory selection changed ---
$cmbSourceDirectory.Add_SelectionChanged({ Update-ConnectionVisibility })
$cmbTargetDirectory.Add_SelectionChanged({ Update-ConnectionVisibility })
#endregion

#region --- Event: Add identity ---
$btnAddDevice.Add_Click({
    Add-Identity -IdentityName $txtIdentity.Text
    $txtIdentity.Clear()
    $txtIdentity.Focus()
})

$txtIdentity.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) {
        Add-Identity -IdentityName $txtIdentity.Text
        $txtIdentity.Clear()
    }
})

$txtIdentity.Add_TextChanged({
    $txtIdentityPH.Visibility = if ($txtIdentity.Text.Length -eq 0) {
        [System.Windows.Visibility]::Visible
    } else {
        [System.Windows.Visibility]::Hidden
    }
})

$txtWaveName.Add_TextChanged({
    $txtWaveNamePH.Visibility = if ($txtWaveName.Text.Length -eq 0) {
        [System.Windows.Visibility]::Visible
    } else {
        [System.Windows.Visibility]::Hidden
    }
})

$txtADDomain.Add_TextChanged({
    $txtADDomainPH.Visibility = if ($txtADDomain.Text.Length -eq 0) {
        [System.Windows.Visibility]::Visible
    } else {
        [System.Windows.Visibility]::Hidden
    }
})
if ($txtADDomain.Text.Length -gt 0) { $txtADDomainPH.Visibility = [System.Windows.Visibility]::Hidden }
#endregion

#region --- Event: Load CSV ---
$btnLoadCsv.Add_Click({
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter   = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.Title    = "Load identity list"
    if ($dlg.ShowDialog() -eq $true) {
        try {
            $rows = Import-Csv -Path $dlg.FileName -ErrorAction Stop
            $col  = $IdentityColumn

            if ($rows.Count -gt 0 -and (-not $col -or -not ($rows[0].PSObject.Properties.Name -contains $col))) {
                $newCol = $rows[0].PSObject.Properties.Name | Select-Object -First 1
                if ($col) { Write-Log "CSV column '$col' not found, using '$newCol'" -Level WARN }
                $col = $newCol
            }

            $added = 0
            foreach ($row in $rows) {
                $name = $row.$col
                if ($name) { Add-Identity -IdentityName $name; $added++ }
            }
            Update-Status "Loaded $added row(s) from CSV"
            Write-Log "CSV loaded: $($dlg.FileName) ($added rows, column '$col')"
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
        "Remove all rows from the list?",
        "Clear all",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question)
    if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
        $identityList.Clear()
        Update-Status "List cleared"
        Write-Log "Identity list cleared"
    }
})
#endregion

#region --- Event: Selection helpers ---
$btnCheckAll.Add_Click({
    foreach ($item in $identityList) { $item.IsSelected = $true }
    $dgResults.Items.Refresh()
    Update-Status "All rows selected"
})

$btnUncheckAll.Add_Click({
    foreach ($item in $identityList) { $item.IsSelected = $false }
    $dgResults.Items.Refresh()
    Update-Status "All rows deselected"
})

$btnUncheckNotFound.Add_Click({
    $count = 0
    foreach ($item in $identityList) {
        if ($item.StatusColor -in @('WARN','ERROR')) {
            $item.IsSelected = $false
            $count++
        }
    }
    $dgResults.Items.Refresh()
    Update-Status "Unchecked $count unresolved row(s)"
    Write-Log "Uncheck not found: $count rows unchecked"
})
#endregion

#region --- Event: Context menu (right-click on grid row) ---
# NOTE: A ContextMenu parsed via XamlReader.Load is not attached to the
# logical tree until it is actually opened, so its child MenuItems do not
# reliably share a NameScope with the rest of the window - $ctx.FindName(...)
# can return $null even though the items exist (confirmed WPF behavior).
# Indexing into $ctx.Items instead of FindName sidesteps NameScope lookup
# entirely and is reliable. Item order matches the XAML:
#   Items[0] = ctxRemoveDevice
#   Items[1] = Separator
#   Items[2] = ctxCheckSelected
#   Items[3] = ctxUncheckSelected
$window.Add_Loaded({
    Update-ConnectionVisibility

    $ctx = $dgResults.ContextMenu

    $miRemoveDevice    = $ctx.Items[0]
    $miCheckSelected   = $ctx.Items[2]
    $miUncheckSelected = $ctx.Items[3]

    $miRemoveDevice.Add_Click({
        $toRemove = @($dgResults.SelectedItems | ForEach-Object { $_ })
        if (-not $toRemove) { return }
        foreach ($item in $toRemove) {
            $identityList.Remove($item) | Out-Null
            Write-Log "Row removed: $($item.Name)"
        }
        Update-Status "Removed $($toRemove.Count) row(s)"
    })

    $miCheckSelected.Add_Click({
        foreach ($item in @($dgResults.SelectedItems)) { $item.IsSelected = $true }
        $dgResults.Items.Refresh()
    })

    $miUncheckSelected.Add_Click({
        foreach ($item in @($dgResults.SelectedItems)) { $item.IsSelected = $false }
        $dgResults.Items.Refresh()
    })
})
#endregion

#region --- Event: Connect ---
$btnConnect.Add_Click({
    $srcDir = Get-DirectoryTag -ComboBox $cmbSourceDirectory
    $tgtDir = Get-DirectoryTag -ComboBox $cmbTargetDirectory
    $needEntra = ($srcDir -eq 'Entra' -or $tgtDir -eq 'Entra')
    $needAD    = ($srcDir -eq 'AD'    -or $tgtDir -eq 'AD')

    if (-not $needEntra -and -not $needAD) {
        [System.Windows.MessageBox]::Show("Select a Source directory (and optionally a Target directory) first.",
            "Nothing to connect", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    if ($needEntra) {
        $reqMods = @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Users')
        $missing = Test-BCGraphModules -Required $reqMods
        if ($missing.Count -gt 0) {
            $msg = "The following required Graph modules are not installed:`n`n$($missing -join "`n")`n`nInstall now from PowerShell Gallery (current user scope)?"
            $confirm = [System.Windows.MessageBox]::Show($msg, "Missing modules",
                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
            if ($confirm -eq [System.Windows.MessageBoxResult]::Yes) {
                Update-Status "Installing required Graph modules..."
                foreach ($m in $missing) {
                    try {
                        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                        Write-Log "Installed module: $m"
                    } catch {
                        [System.Windows.MessageBox]::Show("Failed to install $m`:`n$($_.Exception.Message)",
                            "Install failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                        Write-Log "Module install failed: $m - $($_.Exception.Message)" -Level ERROR
                    }
                }
            }
        }

        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
            Import-Module Microsoft.Graph.Users -ErrorAction Stop
            Update-Status "Signing in to Entra ID (check for a browser window)..."
            Write-Log "Connect-MgGraph requested (Device.Read.All, User.Read.All)"
            Connect-MgGraph -Scopes @('Device.Read.All','User.Read.All') -NoWelcome -ErrorAction Stop
            $ctx = Get-MgContext
            $script:GraphConnected = $true
            $lblGraphStatus.Text = "Entra ID: Connected as $($ctx.Account)"
            $lblGraphStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
            Write-Log "Connected to Microsoft Graph as $($ctx.Account), tenant $($ctx.TenantId)"
        } catch {
            $script:GraphConnected = $false
            $lblGraphStatus.Text = "Entra ID: Not connected"
            $lblGraphStatus.Foreground = [System.Windows.Media.Brushes]::Firebrick
            [System.Windows.MessageBox]::Show("Entra ID sign-in failed:`n$($_.Exception.Message)",
                "Connect-MgGraph failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            Write-Log "Connect-MgGraph failed: $($_.Exception.Message)" -Level ERROR
        }
    }

    if ($needAD) {
        $domain = $txtADDomain.Text.Trim()
        if (-not $domain) {
            [System.Windows.MessageBox]::Show("Enter an AD Domain before connecting to Active Directory.",
                "AD Domain required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            $txtADDomain.Focus()
        } elseif (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $script:ADAvailable = $false
            $lblADStatus.Text = "Active Directory: Module not installed"
            $lblADStatus.Foreground = [System.Windows.Media.Brushes]::Firebrick
            Write-Log "ActiveDirectory module not found on this machine" -Level WARN
            Show-BCRsatInstallDialog -Owner $window
        } else {
            try {
                Update-Status "Connecting to Active Directory ($domain)..."
                Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
                Get-ADDomain -Server $domain -ErrorAction Stop | Out-Null
                $script:ADAvailable = $true
                $lblADStatus.Text = "Active Directory: Connected to $domain"
                $lblADStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
                Write-Log "Active Directory connected: $domain"
            } catch {
                $script:ADAvailable = $false
                $lblADStatus.Text = "Active Directory: Not connected"
                $lblADStatus.Foreground = [System.Windows.Media.Brushes]::Firebrick
                [System.Windows.MessageBox]::Show("Active Directory connection failed:`n$($_.Exception.Message)",
                    "AD connection failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                Write-Log "AD connection failed ($domain): $($_.Exception.Message)" -Level ERROR
            }
        }
    }

    Update-Status "Connection check complete"
})
#endregion

#region --- Helper: resolve one side (Source or Target) for one identity ---
function Resolve-OneSide {
    param(
        [string]$Mode,
        [string]$Directory,
        [string]$Domain,
        [string]$Name,
        [string]$Label
    )

    $result = [PSCustomObject]@{
        ObjectId          = ''
        UserPrincipalName = ''
        SamAccountName    = ''
        Note              = ''
    }

    if ($Directory -eq 'None') { return $result }

    $escaped = $Name.Replace("'", "''")

    try {
        if ($Mode -eq 'Devices') {
            if ($Directory -eq 'Entra') {
                $found = @(Get-MgDevice -Filter "displayName eq '$escaped'" -Property Id,DisplayName -ErrorAction Stop)
                if ($found.Count -eq 0) {
                    $result.Note = "${Label}: not found in Entra ID"
                } elseif ($found.Count -eq 1) {
                    $result.ObjectId = $found[0].Id
                    $result.Note = "${Label}: found (Entra ID)"
                } else {
                    $result.Note = "${Label}: ambiguous ($($found.Count) matches in Entra ID)"
                }
            } else {
                $found = @(Get-ADComputer -Filter "Name -eq '$escaped'" -Server $Domain -Properties ObjectGUID,SamAccountName -ErrorAction Stop)
                if ($found.Count -eq 0) {
                    $result.Note = "${Label}: not found in AD ($Domain)"
                } elseif ($found.Count -eq 1) {
                    $result.ObjectId = $found[0].ObjectGUID.ToString()
                    # AD computer SamAccountName always ends in a literal trailing "$"
                    # (e.g. "COMPUTER01$") - stripped here per requested CSV format.
                    $result.SamAccountName = $found[0].SamAccountName -replace '\$$', ''
                    $result.Note = "${Label}: found (AD $Domain)"
                } else {
                    $result.Note = "${Label}: ambiguous ($($found.Count) matches in AD)"
                }
            }
        } else {
            if ($Directory -eq 'Entra') {
                try {
                    $u = Get-MgUser -UserId $Name -Property Id,UserPrincipalName -ErrorAction Stop
                    $result.ObjectId = $u.Id
                    $result.UserPrincipalName = $u.UserPrincipalName
                    # SamAccountName is intentionally left blank here - it should only be
                    # populated from an actual Active Directory lookup, never from Entra
                    # (even for hybrid-synced users where OnPremisesSamAccountName exists).
                    $result.Note = "${Label}: found (Entra ID)"
                } catch {
                    $result.Note = "${Label}: not found in Entra ID"
                }
            } else {
                $found = @(Get-ADUser -Filter "UserPrincipalName -eq '$escaped'" -Server $Domain -Properties ObjectGUID,SamAccountName,UserPrincipalName -ErrorAction Stop)
                if ($found.Count -eq 0) {
                    $result.Note = "${Label}: not found in AD ($Domain)"
                } elseif ($found.Count -eq 1) {
                    $result.ObjectId = $found[0].ObjectGUID.ToString()
                    $result.UserPrincipalName = $found[0].UserPrincipalName
                    $result.SamAccountName = $found[0].SamAccountName
                    $result.Note = "${Label}: found (AD $Domain)"
                } else {
                    $result.Note = "${Label}: ambiguous ($($found.Count) matches in AD)"
                }
            }
        }
    } catch {
        $result.Note = "${Label}: error - $($_.Exception.Message)"
    }

    return $result
}
#endregion

#region --- Event: Lookup Identities ---
$btnLookup.Add_Click({
    $mode   = Get-CurrentMode
    $srcDir = Get-DirectoryTag -ComboBox $cmbSourceDirectory
    $tgtDir = Get-DirectoryTag -ComboBox $cmbTargetDirectory
    $domain = $txtADDomain.Text.Trim()

    if (($srcDir -eq 'Entra' -or $tgtDir -eq 'Entra') -and -not $script:GraphConnected) {
        [System.Windows.MessageBox]::Show("Connect to Entra ID first (click Connect).", "Not connected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    if (($srcDir -eq 'AD' -or $tgtDir -eq 'AD') -and -not $script:ADAvailable) {
        [System.Windows.MessageBox]::Show("Connect to Active Directory first (click Connect).", "Not connected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    if (($srcDir -eq 'AD' -or $tgtDir -eq 'AD') -and -not $domain) {
        [System.Windows.MessageBox]::Show("Enter an AD Domain before running a lookup.", "AD Domain required",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $targets = Get-SelectedIdentities
    if (-not $targets) {
        [System.Windows.MessageBox]::Show("No rows selected.", "Nothing selected",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    Update-Status "Looking up $($targets.Count) row(s)..."
    Write-Log "Lookup started: mode=$mode source=$srcDir target=$tgtDir domain=$domain rows=$($targets.Count)"
    $btnLookup.IsEnabled = $false

    foreach ($item in $targets) {
        $name = $item.Name

        $src = Resolve-OneSide -Mode $mode -Directory $srcDir -Domain $domain -Name $name -Label 'Source'
        if ($src.ObjectId)          { $item.SourceObjectId = $src.ObjectId }
        if ($src.UserPrincipalName) { $item.SourceUserPrincipalName = $src.UserPrincipalName }
        if ($src.SamAccountName)    { $item.SourceSamAccountName = $src.SamAccountName }

        $tgtNote = ''
        if ($tgtDir -ne 'None') {
            $tgt = Resolve-OneSide -Mode $mode -Directory $tgtDir -Domain $domain -Name $name -Label 'Target'
            if ($tgt.ObjectId)          { $item.TargetObjectId = $tgt.ObjectId }
            if ($tgt.UserPrincipalName) { $item.TargetUserPrincipalName = $tgt.UserPrincipalName }
            if ($tgt.SamAccountName)    { $item.TargetSamAccountName = $tgt.SamAccountName }
            $tgtNote = $tgt.Note
        }

        $item.LookupStatus = if ($tgtNote) { "$($src.Note) | $tgtNote" } else { $src.Note }

        if (-not $item.Comments) {
            if ($src.Note -notlike 'Source: found*') { $item.Comments = $src.Note }
            elseif ($tgtNote -and $tgtNote -notlike 'Target: found*') { $item.Comments = $tgtNote }
        }

        $allNotes = "$($src.Note) $tgtNote"
        if ($allNotes -like '*error*') { $item.StatusColor = 'ERROR' }
        elseif ($allNotes -like '*not found*') { $item.StatusColor = 'ERROR' }
        elseif ($allNotes -like '*ambiguous*') { $item.StatusColor = 'WARN' }
        else { $item.StatusColor = 'OK' }

        Write-Log "${name}: $($item.LookupStatus)"
        $dgResults.Items.Refresh()
        $lblStatus.Text = "Looked up: $name"
        $dgResults.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
    }

    $btnLookup.IsEnabled = $true
    Update-Status "Lookup complete"
    Write-Log "Lookup complete"
})
#endregion

#region --- Event: Export Wave CSV ---
$btnExportCsv.Add_Click({
    $wave = $txtWaveName.Text.Trim()
    if (-not $wave) {
        [System.Windows.MessageBox]::Show("Wave Name is required before exporting.", "Wave Name required",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        $txtWaveName.Focus()
        return
    }
    if ($identityList.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No rows in the list to export.", "Nothing to export",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }

    $mode = Get-CurrentMode
    $modeSuffix = if ($mode -eq 'Users') { 'USERS' } else { 'DEVICES' }
    $safeWave = ($wave -replace '[\\/:*?"<>|]', '_')
    $defaultName = "${safeWave}_${modeSuffix}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter   = "CSV files (*.csv)|*.csv"
    $dlg.FileName = $defaultName
    if ($OutputRoot -and (Test-Path $OutputRoot)) { $dlg.InitialDirectory = $OutputRoot }

    if ($dlg.ShowDialog() -eq $true) {
        # Column set and exact header names differ by mode, per spec:
        #   Users:   SourceObjectId,SourceUserPrincipalName (Optional),SourceSamAccountName (Optional),
        #            TargetObjectId (Optional),TargetUserPrincipalName (Optional),TargetSamAccountName (Optional),Comments (Optional)
        #   Devices: SourceObjectId,SourceSamAccountName (Optional),Comments (Optional)
        # NOTE: Devices intentionally drops SourceUserPrincipalName and all Target
        # fields from the export. If a Target directory was selected and resolved
        # for devices, that data still shows in the grid but is NOT written to the
        # CSV - only Source ObjectId, Source SamAccountName, and Comments are.
        if ($mode -eq 'Users') {
            $rows = foreach ($item in $identityList) {
                [PSCustomObject][ordered]@{
                    'SourceObjectId'                    = $item.SourceObjectId
                    'SourceUserPrincipalName (Optional)' = $item.SourceUserPrincipalName
                    'SourceSamAccountName (Optional)'   = $item.SourceSamAccountName
                    'TargetObjectId (Optional)'         = $item.TargetObjectId
                    'TargetUserPrincipalName (Optional)' = $item.TargetUserPrincipalName
                    'TargetSamAccountName (Optional)'   = $item.TargetSamAccountName
                    'Comments (Optional)'               = $item.Comments
                }
            }
        } else {
            $rows = foreach ($item in $identityList) {
                [PSCustomObject][ordered]@{
                    'SourceObjectId'                  = $item.SourceObjectId
                    'SourceSamAccountName (Optional)' = $item.SourceSamAccountName
                    'Comments (Optional)'             = $item.Comments
                }
            }
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8 -UseQuotes AsNeeded
        Update-Status "Exported $($rows.Count) row(s) to $($dlg.FileName)"
        Write-Log "Exported wave CSV '$wave' (mode=$mode): $($dlg.FileName) ($($rows.Count) rows)"

        $unresolved = @($identityList | Where-Object { $_.StatusColor -ne 'OK' }).Count
        $note = if ($unresolved -gt 0) { "`n`n$unresolved row(s) are not fully resolved (not found, ambiguous, error, or not yet looked up)." } else { '' }
        [System.Windows.MessageBox]::Show("Exported $($rows.Count) row(s) to:`n$($dlg.FileName)$note",
            "Export complete", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
})
#endregion

#region --- Event: Show log ---
$btnShowLog.Add_Click({
    $lf = $LogFile
    Start-Process pwsh -ArgumentList @(
        '-NoExit',
        '-Command',
        "Write-Host 'BriComp Identity Object Mapper - Live Log' -ForegroundColor Cyan; Write-Host 'File: $lf' -ForegroundColor DarkGray; Write-Host ''; Get-Content -Path '$lf' -Wait -Tail 40"
    )
    Write-Log "Log viewer opened"
})
#endregion

#region --- Load CSV on startup if provided ---
if ($CsvPath -and (Test-Path $CsvPath)) {
    try {
        $rows = Import-Csv -Path $CsvPath -ErrorAction Stop
        $col = $IdentityColumn
        if ($rows.Count -gt 0 -and (-not $col -or -not ($rows[0].PSObject.Properties.Name -contains $col))) {
            $col = $rows[0].PSObject.Properties.Name | Select-Object -First 1
        }
        $added = 0
        foreach ($row in $rows) {
            $name = $row.$col
            if ($name) { Add-Identity -IdentityName $name; $added++ }
        }
        Update-Status "Loaded $added row(s) from $CsvPath"
        Write-Log "Startup CSV loaded: $CsvPath ($added rows, column '$col')"
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQfyybgGIlnaSw
# hjA8bjXOSLhdYqozIEr8lyRgKt1vj6CCIWswggWNMIIEdaADAgECAhAOmxiO+dAt
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
# AwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# MB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEy
# NTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYgMTCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz/TjGMPh8CQ6RucZC
# Lv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMxt9l3Dm5oCG5Z69Aq
# EN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkUxDjLJhBI+qiZP3ai
# xAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZEeoy/0B8Zm+nzKBgS
# WG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEOvZ3fP2VxIgeR28Y9
# 0ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn/FnbuTegcoRcFEZ9
# FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolvykE1jive4PuESaT8
# zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZCZVl6pNUOCybMc/k
# xC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW83vV8d67WowuUs0Hd
# Y4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjdzi0xlp70W+ow63/0
# ++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9icflf55W1zOEKAcJVA
# IXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBQUyWOK
# MC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezL
# TjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsG
# AQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNy
# dDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGln
# aUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQIENoyKqxrFbJblzCV
# P/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGhz45F4Wt4WGdNdIbR
# Wt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfHhHv1mx9EVh71SJel
# cEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA02nXD1nYLI7Ijmiv
# +xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7KVyF5aQyWP4KhVGUU
# Tnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jhtKVnd3D0rbHXEu27o
# +Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvKqzpAec9etE14jOQA
# zLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs0NzU/+xX7NbEdcof
# y0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGRsMNbXmMREUTgu+Oc
# a7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJNt68Yfim6cjEzVSi
# LSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yWQPT9gzCCB3kwggVh
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
# BDEiBCABR/GyUixBG7XFBwujw9qml6jQhFI6D09VFx4eqyN7ajANBgkqhkiG9w0B
# AQEFAASCAgDYhDClz3T8N7RIGzc5AOXuDs3GkehRGlrs89fvUUWfuYKbqFLPqZ/d
# 6FPxRICFs9WyJDiTlOUvOOWWz2Du4aSiU5/OCBCM+7/XwSD3VQivvM7n+PEZ7hsr
# gygrgmj+sIKqT7v4rq+kC3t2vhg0M0MY6rtwSFkzXbPCf0IsIAwlWis83vqeQI9s
# gF0RlLfzzc0dT68h8SIzHyLOOCk5fF0wI3K1hZ1Qkzq4wmIC6FdNvgVLpp1f8KEd
# QP3+50Mcu+vnjQTp79yzp1llQZqoXEmnJzhxbl+PF8HrowsrqPwpjTkfuLdE3iBf
# TrvzLeCZt+U8DFavfiMkVXR86FLRUHgDq6ZNgO1u99ztLVy5/u5JJxmTwU+hTdjK
# b+mu0B/m3zh1lC1ljr0rHvkx28w0zZ5ObL8XXf7U1V4LWUCVZsc/hgYPJbzR+Gud
# B+8Zu7QNFXi8M4KtLPzXPCx03rTRqXwb/tHBY5ZdvRm2ebyadwt8WkEMUB1+ph9v
# ewsBwKyS1lJ8YPJLwwayNHPdlCKuxFMZFRBaDVleoMw14KfE1KyyuWbO8eNwZVSy
# SnKpvW6po+HJdBatshqG3e/38IqdGmvoASGSvplgCWkx87YtlxFK8vdbpeYhWVqC
# GZDiO8Sk6q7oalTKONLmFaXjuA2/LLboNkSMnJJanKri65FGBnWZ1qGCAyYwggMi
# BgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQK
# Ew5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBU
# aW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTECEAhP3DNPfkVO28MP
# j/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcB
# MBwGCSqGSIb3DQEJBTEPFw0yNjA5MTgxNjE3MDdaMC8GCSqGSIb3DQEJBDEiBCBn
# Zxu1oo98DWWOGJZwBL2y62PcrnDTkfXgqlqJ7/J2ejANBgkqhkiG9w0BAQEFAASC
# AgAOThga8X37+5l6Z4ADmNydsWWjQ5225ESR6nzxgo8IPr/ehd4ISAfTRCLO6akr
# kmgX3jXuY3NCSvqybo+Ox46NguLK10+h7vIT59KnBbSakMHt9r4H1f4ZdHhKV+Ou
# rcioA45SU4ZCVpqXr4fAYsdCPOTJHa6t65Omeu9sT426Tf1TdcLQDDCdVnHF/cMC
# tZQuyQ18ah/qx1lasQx/c5GwgN3XW8pTiTTGwc79TRPpT4v51oVJJkPfU/d2zBlu
# vJB9vmmd+8NjoI+yIm3yudYzZs5VnXH3qxUZC6stSJQ0XEU8raPyuBHNkQ/4cuEX
# 0d5/KN+E48g8WP+A+tMiLVHCwtt9ew9FXoLO1IIKcE2e+RsEKstxjlqGpU6wiY8S
# G69+3yxWbipZ3EQCHDDyL6Znx2v8y4wABK0QqU5+XjU72ZqhfMCs9VNAc/Pvq5oo
# 9ybEWW8KRek+pHDBATcQW/XiX6OfZafd2c2/idsvs5vKmSUV6fu4Y6GZecCHdAP/
# 3LsuUEDPmagaCQC3LQGztplsY3MdPO7hK44eh4DBJIr7/jg0H+GiuMqBoOpDWsrx
# dffJ8tSVROWPla44waT407XLERPNYxze33VlbVgqP0cjpG15rnvxBHbOib94MQsF
# lT/t6e5ufDSy1Vnlq0VU2EJBOS8o2sj37du0Vl05Y9va8g==
# SIG # End signature block
