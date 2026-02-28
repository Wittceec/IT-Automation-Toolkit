<# 
   =============================================================================
   EMAIL SUITE GUI (WPF/XAML) v8.1 - Patch Dates Fix
   Author: Chris Wittman
   =============================================================================
#>

# 1. Load Assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 2. Path Configuration
$Root = $PSScriptRoot
$TplDir = Join-Path $Root 'templates'
$LogDir = Join-Path $Root 'logs'
$DLists = Join-Path $Root 'dlists'
$PatchesExcel = Join-Path $Root "Patches.xlsx"
$ChangesExcel = Join-Path $Root "ChangeRequests.xlsx"

# Sibling Folder Logic
$ParentScripts = Split-Path $Root -Parent
$GrandParent   = Split-Path $ParentScripts -Parent
$SchedulesDir  = Join-Path $GrandParent "Data Center Operations - Operations Schedules"

# Global State
$Global:ParsedData = @{}
$Global:PasteOutput = ""
$Global:PromptOutput = ""
$Global:DlSelection = @()

# 3. The XAML UI Definition
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Data Center Operations Suite v8.1" Height="800" Width="1100" 
    WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
    Background="#121212" Foreground="#E0E0E0" FontFamily="Segoe UI">

    <Window.Resources>
        <SolidColorBrush x:Key="BgDark" Color="#121212"/>
        <SolidColorBrush x:Key="BgCard" Color="#1E1E1E"/>
        <SolidColorBrush x:Key="BgInput" Color="#2D2D30"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#00E676"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#2979FF"/>
        <SolidColorBrush x:Key="TextDim" Color="#AAAAAA"/>

        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#444"/>
        </Style>

        <Style TargetType="TabItem">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="20,12"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#888"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Grid>
                            <Border Name="Border" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,3" Margin="0,0,5,0">
                                <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="15,10"/>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource AccentGreen}" />
                                <Setter Property="Foreground" Value="{StaticResource AccentGreen}" />
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#FFFFFF" />
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#333"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Name="Border" Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#444"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource AccentGreen}"/>
                                <Setter Property="Foreground" Value="#000"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="BtnPrimary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#1B5E20"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#2E7D32"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{StaticResource BgCard}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Foreground" Value="#DDD"/>
        </Style>
        <Style TargetType="ListBoxItem">
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border Name="Border" Background="Transparent" CornerRadius="4" Margin="2">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#333"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentGreen}"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="#252525"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{StaticResource BgInput}"/>
            <Setter Property="Foreground" Value="#AAA"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/>
            <Setter Property="BorderBrush" Value="#222"/>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Foreground" Value="#EEE"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Padding" Value="5,8"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
            <Setter Property="BorderBrush" Value="#222"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#2E352E"/>
                    <Setter Property="Foreground" Value="{StaticResource AccentGreen}"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#252526"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="GroupBox">
            <Setter Property="BorderBrush" Value="#333"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="HeaderTemplate">
                <Setter.Value>
                    <DataTemplate>
                        <TextBlock Text="{Binding}" FontWeight="Bold" Foreground="{StaticResource AccentBlue}" Margin="5,0,5,0"/>
                    </DataTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <RowDefinition Height="*"/>    <RowDefinition Height="Auto"/> <RowDefinition Height="150"/>  </Grid.RowDefinitions>

        <Border Grid.Row="0" BorderBrush="#333" BorderThickness="0,0,0,1" Margin="0,0,0,15" Padding="0,0,0,10">
            <StackPanel Orientation="Horizontal">
                <TextBlock Text="DC OPS" FontWeight="Bold" FontSize="24" Foreground="{StaticResource AccentGreen}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Text="// AUTOMATION SUITE" FontWeight="Light" FontSize="24" Foreground="{StaticResource TextDim}" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>

        <TabControl Grid.Row="1" Background="Transparent" BorderThickness="0">
            
            <TabItem Header="PATCHING">
                <Grid Margin="0,15,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
                        <Button Name="btnLoadPatches" Content="Reload Excel Data" Width="160"/>
                        <TextBlock Text="Select a maintenance window below to generate notification." VerticalAlignment="Center" Foreground="#666" Margin="15,0" FontStyle="Italic"/>
                    </StackPanel>

                    <Border Grid.Row="1" Background="{StaticResource BgCard}" CornerRadius="5" Padding="2">
                        <ListView Name="lstPatches" Background="Transparent" BorderThickness="0">
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="SYSTEM" Width="200" DisplayMemberBinding="{Binding System}"/>
                                    <GridViewColumn Header="ENVIRONMENT" Width="140" DisplayMemberBinding="{Binding Env}"/>
                                    <GridViewColumn Header="DATE" Width="180" DisplayMemberBinding="{Binding DisplayDate}"/>
                                    <GridViewColumn Header="TIME" Width="120" DisplayMemberBinding="{Binding DisplayTime}"/>
                                    <GridViewColumn Header="STATUS" Width="120" DisplayMemberBinding="{Binding Status}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </Border>

                    <StackPanel Orientation="Horizontal" Grid.Row="2" HorizontalAlignment="Right" Margin="0,15,0,0">
                        <Button Name="btnPatchBefore" Content="Generate BEFORE Email" Width="200" Style="{StaticResource BtnPrimary}"/>
                        <Button Name="btnPatchAfter" Content="Generate AFTER Email" Width="200"/>
                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="SERVICE ALERTS">
                <Grid Margin="0,15,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="300"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <DockPanel Grid.Column="0" Margin="0,0,20,0">
                        <TextBlock Text="AVAILABLE TEMPLATES" DockPanel.Dock="Top" Foreground="{StaticResource AccentBlue}" FontWeight="Bold" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource BgCard}" CornerRadius="5" Padding="2">
                            <ListBox Name="lstTemplates" DisplayMemberPath="Name"/>
                        </Border>
                    </DockPanel>

                    <StackPanel Grid.Column="1">
                        
                        <GroupBox Header="1. SOURCE DATA (AUTOFILL)" Padding="15" Margin="0,0,0,20">
                            <StackPanel>
                                <TextBlock Text="Paste the raw email or load a file to automatically map fields." Foreground="#888" Margin="0,0,0,10"/>
                                
                                <StackPanel Orientation="Horizontal">
                                    <Button Name="btnPasteText" Content="Paste Text" Width="140" Style="{StaticResource BtnPrimary}"/>
                                    <Button Name="btnLoadFile" Content="Load File" Width="140"/>
                                    <Button Name="btnClearAuto" Content="Clear" Width="80" Background="#B71C1C"/>
                                </StackPanel>
                                
                                <Border Background="{StaticResource BgInput}" CornerRadius="4" Padding="10" Margin="0,10,0,0">
                                    <StackPanel>
                                        <DockPanel LastChildFill="True">
                                            <TextBlock Text="STATUS:" FontWeight="Bold" Foreground="#666" Margin="0,0,10,0"/>
                                            <TextBlock Name="txtAutoStatus" Text="No data loaded. Template will prompt for input." Foreground="#888" TextTrimming="CharacterEllipsis"/>
                                        </DockPanel>
                                        <TextBox Name="txtParsedPreview" Height="80" Margin="0,10,0,0" Background="#111" Foreground="{StaticResource AccentGreen}" 
                                                 FontFamily="Consolas" BorderThickness="0" IsReadOnly="True" Visibility="Collapsed" VerticalScrollBarVisibility="Auto"/>
                                    </StackPanel>
                                </Border>
                            </StackPanel>
                        </GroupBox>

                        <TextBlock Text="2. GENERATE DRAFT" Foreground="{StaticResource AccentGreen}" FontWeight="Bold" Margin="0,0,0,5"/>
                        <Button Name="btnRunTemplate" Content="LAUNCH TEMPLATE" HorizontalAlignment="Left" Width="250" Height="45" FontSize="14" FontWeight="Bold" Style="{StaticResource BtnPrimary}"/>
                        <TextBlock Text="(Includes Distribution List Selection)" Foreground="#555" Margin="5,5,0,0" FontSize="11" FontStyle="Italic"/>

                    </StackPanel>
                </Grid>
            </TabItem>

            <TabItem Header="TOOLS &amp; ADMIN">
                <WrapPanel Margin="0,15,0,0">
                    <Border Background="{StaticResource BgCard}" CornerRadius="5" Width="300" Margin="0,0,20,20" Padding="15">
                        <StackPanel>
                            <TextBlock Text="SCHEDULES" FontWeight="Bold" Foreground="{StaticResource AccentBlue}" Margin="0,0,0,10"/>
                            <TextBlock Text="Print all schedules found in 'Operations Schedules'." Foreground="#888" TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Button Name="btnPrintSchedules" Content="Print All Schedules" HorizontalAlignment="Stretch"/>
                        </StackPanel>
                    </Border>

                    <Border Background="{StaticResource BgCard}" CornerRadius="5" Width="300" Margin="0,0,20,20" Padding="15">
                        <StackPanel>
                            <TextBlock Text="PATCH MAINTENANCE" FontWeight="Bold" Foreground="{StaticResource AccentBlue}" Margin="0,0,0,10"/>
                            <TextBlock Text="Update Excel dates based on frequency rules." Foreground="#888" TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                <TextBox Name="txtYear" Width="60" Margin="2" Background="{StaticResource BgInput}" Foreground="#FFF" TextAlignment="Center" Padding="5"/>
                                <TextBox Name="txtMonth" Width="40" Margin="2" Background="{StaticResource BgInput}" Foreground="#FFF" TextAlignment="Center" Padding="5"/>
                            </StackPanel>
                            <Button Name="btnUpdateDates" Content="Run Update" HorizontalAlignment="Stretch" Margin="0,10,0,0"/>
                        </StackPanel>
                    </Border>

                    <Border Background="{StaticResource BgCard}" CornerRadius="5" Width="300" Margin="0,0,20,20" Padding="15">
                        <StackPanel>
                            <TextBlock Text="LOCAL CLEANUP" FontWeight="Bold" Foreground="{StaticResource AccentBlue}" Margin="0,0,0,10"/>
                            <TextBlock Text="Clean .rdp files from Downloads folder." Foreground="#888" TextWrapping="Wrap" Margin="0,0,0,15"/>
                            <Button Name="btnCleanRDP" Content="Clean RDP Files" HorizontalAlignment="Stretch"/>
                        </StackPanel>
                    </Border>
                </WrapPanel>
            </TabItem>

        </TabControl>

        <TextBlock Grid.Row="2" Text="SYSTEM LOG" Foreground="#555" FontWeight="Bold" FontSize="11" Margin="5,10,0,2"/>
        <Border Grid.Row="3" Background="#000" BorderBrush="#333" BorderThickness="1" CornerRadius="3" Padding="5">
            <TextBox Name="txtLog" Background="Transparent" Foreground="{StaticResource AccentGreen}" FontFamily="Consolas" FontSize="12"
                     BorderThickness="0" IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"/>
        </Border>

    </Grid>
</Window>
"@

# 4. Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try { $window = [Windows.Markup.XamlReader]::Load($reader) } catch { Write-Host "XAML Error: $_"; exit }

# 5. Connect UI Elements
$btnLoadPatches   = $window.FindName("btnLoadPatches")
$lstPatches       = $window.FindName("lstPatches")
$btnPatchBefore   = $window.FindName("btnPatchBefore")
$btnPatchAfter    = $window.FindName("btnPatchAfter")
$lstTemplates     = $window.FindName("lstTemplates")
$btnRunTemplate   = $window.FindName("btnRunTemplate")
$btnLoadFile      = $window.FindName("btnLoadFile")
$btnPasteText     = $window.FindName("btnPasteText")
$btnClearAuto     = $window.FindName("btnClearAuto")
$txtAutoStatus    = $window.FindName("txtAutoStatus")
$txtParsedPreview = $window.FindName("txtParsedPreview")
$btnPrintSchedules= $window.FindName("btnPrintSchedules")
$btnUpdateDates   = $window.FindName("btnUpdateDates")
$txtYear          = $window.FindName("txtYear")
$txtMonth         = $window.FindName("txtMonth")
$btnCleanRDP      = $window.FindName("btnCleanRDP")
$txtLog           = $window.FindName("txtLog")

# Initialize Date Defaults
$nextMonth = (Get-Date).AddMonths(1)
$txtYear.Text = $nextMonth.Year.ToString()
$txtMonth.Text = $nextMonth.Month.ToString()

# =============================================================================
# LOGGING & HELPERS
# =============================================================================
function Log-Gui {
    param([string]$Msg, [string]$Color="Normal")
    $ts = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$ts] $Msg`n")
    $txtLog.ScrollToEnd()
    [System.Windows.Forms.Application]::DoEvents()
    try { Add-Content -Path (Join-Path $LogDir "emailsuite_gui.log") -Value "[$ts] $Msg" } catch {}
}

function Coalesce { param([Parameter(ValueFromRemainingArguments=$true)]$V) foreach($x in $V){if($x -and "$x".Trim()){return "$x"}}; return "" }
function Fmt-Date { param($v) try{[datetime]$v|Get-Date -F "MMMM d, yyyy"}catch{Coalesce $v} }
function Fmt-Time { param($v) try{([datetime]$v).ToString("h:mmtt").ToLower()}catch{Coalesce $v} }
function Read-Json { param([string]$Path) Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json }
function Safe-Join([object]$v){ if($null -eq $v){""} elseif($v -is [array]){$v -join ';'} else{[string]$v} }

# =============================================================================
# CUSTOM DIALOGS
# =============================================================================

function Show-PasteWindow {
    [xml]$dXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Paste Email Body" Height="450" Width="700" WindowStartupLocation="CenterScreen" Background="#1E1E1E" Foreground="#E0E0E0" Topmost="True" ResizeMode="NoResize">
    <Grid Margin="15">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Text="PASTE RAW TEXT BELOW:" FontWeight="Bold" Foreground="#00E676" Margin="0,0,0,10"/>
        <TextBox Name="txtPaste" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap" Background="#2D2D30" Foreground="#FFF" BorderThickness="0" Padding="5" FontFamily="Consolas"/>
        <StackPanel Orientation="Horizontal" Grid.Row="2" HorizontalAlignment="Right" Margin="0,15,0,0">
            <Button Name="btnCancel" Content="Cancel" Width="100" Margin="5" Background="#333" Foreground="#FFF"/>
            <Button Name="btnOk" Content="PARSE DATA" Width="140" Margin="5" Background="#1B5E20" Foreground="#FFF" IsDefault="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    $r = (New-Object System.Xml.XmlNodeReader $dXaml); $w = [Windows.Markup.XamlReader]::Load($r)
    $txt = $w.FindName("txtPaste"); $btnOk = $w.FindName("btnOk"); $btnCancel = $w.FindName("btnCancel")
    $Global:PasteOutput = ""
    $btnOk.Add_Click({ $Global:PasteOutput = $txt.Text; $w.Close() }); $btnCancel.Add_Click({ $w.Close() })
    $w.ShowDialog() | Out-Null
    return $Global:PasteOutput
}

function Show-PromptWindow {
    param($Title, $Label, $DefaultValue)
    [xml]$dXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="$Title" Height="200" Width="500" WindowStartupLocation="CenterScreen" Background="#1E1E1E" Foreground="#E0E0E0" Topmost="True" ResizeMode="NoResize">
    <Grid Margin="20">
        <StackPanel>
            <TextBlock Text="$Label" FontWeight="Bold" Foreground="#2979FF" Margin="0,0,0,10" FontSize="14"/>
            <TextBox Name="txtInput" Background="#2D2D30" Foreground="#FFF" Padding="8" BorderThickness="0" FontSize="13"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,25,0,0">
                <Button Name="btnOk" Content="OK" Width="100" Background="#1B5E20" Foreground="#FFF" IsDefault="True"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"@
    $r = (New-Object System.Xml.XmlNodeReader $dXaml); $w = [Windows.Markup.XamlReader]::Load($r)
    $txt = $w.FindName("txtInput"); $btnOk = $w.FindName("btnOk")
    $txt.Text = $DefaultValue; $txt.Focus()
    $Global:PromptOutput = $null
    $btnOk.Add_Click({ $Global:PromptOutput = $txt.Text; $w.Close() })
    $w.ShowDialog() | Out-Null
    return $Global:PromptOutput
}

function Show-DListSelector {
    if(-not (Test-Path $DLists)){ New-Item -ItemType Directory -Path $DLists -Force | Out-Null }
    $files = Get-ChildItem -Path $DLists -Filter *.txt
    if($files.Count -eq 0){ return "" }

    [xml]$dXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Distribution Lists" Height="450" Width="400" WindowStartupLocation="CenterScreen" Background="#1E1E1E" Foreground="#E0E0E0" Topmost="True">
    <Grid Margin="15">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Text="ADD TO BCC:" Grid.Row="0" Margin="0,0,0,10" Foreground="#2979FF" FontWeight="Bold"/>
        <ListBox Name="lstDl" Grid.Row="1" Background="#2D2D30" Foreground="#FFF" SelectionMode="Multiple" BorderThickness="0" Padding="5">
            <ListBox.ItemContainerStyle>
                <Style TargetType="ListBoxItem">
                    <Setter Property="Padding" Value="5"/>
                    <Style.Triggers>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#333"/>
                            <Setter Property="Foreground" Value="#00E676"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </ListBox.ItemContainerStyle>
        </ListBox>
        <Button Name="btnOk" Grid.Row="2" Content="CONFIRM SELECTION" Margin="0,15,0,0" Height="40" Background="#1B5E20" Foreground="#FFF" IsDefault="True"/>
    </Grid>
</Window>
"@
    $r = (New-Object System.Xml.XmlNodeReader $dXaml); $w = [Windows.Markup.XamlReader]::Load($r)
    $lst = $w.FindName("lstDl"); $btnOk = $w.FindName("btnOk")
    foreach($f in $files){ $lst.Items.Add($f.Name) | Out-Null }
    $Global:DlSelection = @()
    $btnOk.Add_Click({ foreach($item in $lst.SelectedItems){ $Global:DlSelection += (Join-Path $DLists $item) }; $w.Close() })
    $w.ShowDialog() | Out-Null
    $allEmails = @(); foreach($p in $Global:DlSelection){ $content = Get-Content -Path $p; $allEmails += $content }
    return ($allEmails -join ";")
}

# =============================================================================
# PARSING ENGINE
# =============================================================================
function Normalize-Text{ param([string]$Text) if(-not $Text){return ""}; $t=$Text -replace "\x00","" -replace "[\xA0]", " " -replace "[\u2010\u2011\u2012\u2013\u2014\u2212]", "-" -replace "`r`n","`n" -replace "`r","`n" -replace "[\t]", " " -replace " {2,}", " "; $t }
function Clean-Field{ param([string]$s) if(-not $s){return ""}; $t=($s -split "`r?`n") |? { $_ -notmatch '^(?i)\s*(For urgent issues|For non-urgent issues|Change Coordinator|This message|Do not reply)\b' -and $_ -notmatch '^(?i)\s*Reference\s*:' -and $_ -notmatch '^(?i)\s*Distribution\s*:' }; $t=($t -join "`r`n") -replace "[\t]"," " -replace ' +',' ' -replace '(\r?\n){3,}',"`r`n`r`n"; $t.Trim() }
function New-RegexAlternation{ param([string[]]$Tokens) ($Tokens|% {[Regex]::Escape($_)}) -join '|' }

function Get-Section{
  param([string]$Text,[string[]]$StartTokens,[string[]]$AllTokens,[string[]]$FooterTokens)
  $lines=$Text -split "`r?`n"; $start=-1; $end=$lines.Length
  $startRx='(?i)^\s*(?:' + (New-RegexAlternation $StartTokens) + ')(?:\s*:|\s*$)'
  for($i=0;$i -lt $lines.Length;$i++){ if($lines[$i] -match $startRx){ $start=$i; break } }
  if($start -lt 0){ return "" }
  $allRx= if($AllTokens){'(?i)^\s*(?:' + (New-RegexAlternation $AllTokens) + ')(?:\s*:|\s*$)'} else {$null}
  $footRx= if($FooterTokens){'(?i)^\s*(?:' + (New-RegexAlternation $FooterTokens) + ')'} else {$null}
  for($j=$start+1;$j -lt $lines.Length;$j++){
    if($allRx -and $lines[$j] -match $allRx){ $end=$j; break }
    if($footRx -and $lines[$j] -match $footRx){ $end=$j; break }
    if($lines[$j] -match '^\s*[\u25A0-\u25FF\u2022\u25CF\u25E6\u25CB\u25A1\u2611\-\*]\s'){ $end=$j; break }
  }
  $first=$lines[$start]; $value=""
  if($first -match ':\s*(.*)$'){ $value=$matches[1] }
  $body=@(); if($end -gt ($start+1)){ $body=$lines[($start+1)..($end-1)] }
  (Clean-Field -s (@($value)+$body -join "`r`n"))
}

function Parse-RequestText{
  param([Parameter(Mandatory)][string]$Text)
  $norm=Normalize-Text $Text; $text=$norm -replace '\s+$',''; $lines=$text -split "`n"
  $map=@{
    Title=@('Title','Subject','What','Summary')
    DateLine=@('Date','Date/Time','When','Window','Start','Schedule')
    Locations=@('Location','Locations','Sites','Site','Facility','Facilities','Locations and/or Departments affected')
    Applications=@('Application','Applications','System','Systems','Service','Services','Impacted Systems','Applications and/or Services affected')
    Details=@('Details','Impact','Description','Info','Information','Notes')
    RequiredActions=@('Required Actions','Action Required','Actions Required','User Action','Customer Action')
  }
  $allTokens=($map.Values|% {$_}) | Select-Object -Unique
  $footerStops=@('For urgent issues','For non-urgent issues','Change Coordinator','This message','Do not reply','Reference','Distribution')
  $result=@{}
  foreach($k in $map.Keys){ $val=Get-Section -Text $text -StartTokens $map[$k] -AllTokens $allTokens -FooterTokens $footerStops; if($val){ $result[$k]=$val } }
  if(-not ($result.ContainsKey('DateLine') -and $result['DateLine'])){ foreach($ln in $lines){ if($ln -match '^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*,\s+[A-Za-z]+\s+\d{1,2}'){ $result['DateLine'] = $ln.Trim(); break } } }
  if($norm -match '(?i)Reference\s*:\s*(.+)') { $result['Reference'] = $matches[1].Trim() }
  if($norm -match '(?i)Distribution\s*:\s*(.+)') { $result['Distribution'] = $matches[1].Trim() }
  if(-not $result['Title']){ foreach($ln in $lines){ $l=$ln.Trim(); if($l -and $l -notmatch '^Notification:' -and $l.Length -gt 5){ $result['Title']=$l; break } } }
  if($result['DateLine']){ $result['DateRange'] = $result['DateLine'] }
  if($result['Title']){ $result['Subject'] = $result['Title'] }
  $result
}

function Get-TextFromMsg{ 
    param([string]$Path) 
    $ol=$null
    try{ 
        $ol=New-Object -ComObject Outlook.Application
        $item=$ol.Session.OpenSharedItem($Path)
        $body = $item.Body
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
        return $body
    } catch { return "" } 
    finally { if($ol){[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol)|Out-Null} }
}

function Get-TextFromDocx{ param([string]$Path) try{ $zip=[IO.Compression.ZipFile]::OpenRead($Path); $entry=$zip.Entries|?{$_.FullName -eq 'word/document.xml'}; if($entry){ $sr=New-Object IO.StreamReader($entry.Open()); $xml=$sr.ReadToEnd(); $sr.Close(); $zip.Dispose(); return ($xml -replace '(?is)<w:tab\s*/?>',"`t" -replace '(?is)<w:br\s*/?>',"`r`n" -replace '(?is)<w:br\s*/?>',"`r`n`r`n" -replace '(?is)<.*?>','').Trim() }; $zip.Dispose() }catch{}; return "" }

# =============================================================================
# WORKER PROCESS GENERATOR (ASYNC LAUNCH)
# =============================================================================
function Start-OutlookWorker {
    param($Vars, $JobData)
    $jsonVars = $Vars | ConvertTo-Json -Depth 5; $varsPath = Join-Path $env:TEMP ("EmailSuiteVars_" + [Guid]::NewGuid() + ".json"); Set-Content -Path $varsPath -Value $jsonVars
    $jsonJob = $JobData | ConvertTo-Json -Depth 5; $jobPath = Join-Path $env:TEMP ("EmailSuiteJob_" + [Guid]::NewGuid() + ".json"); Set-Content -Path $jobPath -Value $jsonJob

    $workerScriptPath = Join-Path $env:TEMP ("EmailSuiteWorker_" + [Guid]::NewGuid() + ".ps1")
    $workerCode = @"
    param([string]`$VarsFile, [string]`$JobFile)
    try {
        `$vars = Get-Content -Raw `$VarsFile | ConvertFrom-Json; `$job  = Get-Content -Raw `$JobFile  | ConvertFrom-Json
        function Rep(`$txt){ if(!`$txt){return ""}; `$pattern='(?s)(\{\{|\[\[)(.*?)(\}\}|\]\])'; [regex]::Replace(`$txt,`$pattern,{param(`$m) `$k=(`$m.Groups[2].Value -replace '<[^>]*>','').Trim(); if(`$vars.`$k){[string]`$vars.`$k}else{`$m.Value}}) }
        `$ol = New-Object -ComObject Outlook.Application; `$mail = `$null
        if (`$job.TemplateMsg -and (Test-Path `$job.TemplateMsg)) {
            `$orig = `$ol.Session.OpenSharedItem(`$job.TemplateMsg); `$mail = `$orig.Copy(); `$mail.HTMLBody = Rep `$mail.HTMLBody
        } else {
            `$mail = `$ol.CreateItem(0); if (`$job.BodyIsHtml) { `$mail.HTMLBody = Rep `$job.Body } else { `$mail.Body = Rep `$job.Body }
        }
        `$mail.Subject = Rep `$job.Subject; `$mail.To = `$job.To; `$mail.CC = `$job.Cc; `$mail.BCC = `$job.Bcc; `$mail.Display()
        Remove-Item `$VarsFile -Force; Remove-Item `$JobFile -Force
    } catch { Add-Content -Path (Join-Path `$env:TEMP 'emailsuite_worker_error.log') -Value (`$_.ToString()) }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$ol) | Out-Null
"@
    Set-Content -Path $workerScriptPath -Value $workerCode
    Log-Gui "Dispatching background worker... (GUI Released)"
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$workerScriptPath`" -VarsFile `"$varsPath`" -JobFile `"$jobPath`""
}

# =============================================================================
# MAIN UI LOGIC
# =============================================================================

function Set-AutofillData {
    param($Hashtable, $SourceName)
    $Global:ParsedData = $Hashtable
    if($Hashtable.Count -gt 0){
        $fields = ($Hashtable.Keys -join ", ")
        $txtAutoStatus.Text = "LOADED FIELDS: $fields"
        $txtAutoStatus.Foreground = "#00E676"
        $prev = ""; foreach($k in $Hashtable.Keys){ $prev += "$k : $($Hashtable[$k])`r`n" }
        $txtParsedPreview.Text = $prev
        $txtParsedPreview.Visibility = "Visible"
        Log-Gui "Autofill data parsed ($fields)"
    } else {
        $txtAutoStatus.Text = "Parsed $SourceName but found no fields."
        $txtAutoStatus.Foreground = "#FFEA00"
    }
}

$btnLoadFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog; $dlg.Filter = "Requests (*.msg;*.docx;*.txt)|*.msg;*.docx;*.txt"
    if($dlg.ShowDialog() -eq "OK"){
        $f = $dlg.FileName; Log-Gui "Reading file: $f..."
        $txt = ""
        switch([IO.Path]::GetExtension($f).ToLower()){ 
            '.msg'{ $txt=Get-TextFromMsg $f } 
            '.docx'{$txt=Get-TextFromDocx $f} 
            '.txt'{$txt=Get-Content -Raw $f} 
        }
        if($txt){ Set-AutofillData (Parse-RequestText $txt) ([IO.Path]::GetFileName($f)) }
        else { Log-Gui "Warning: Could not read text from $f" "Error" }
    }
})

$btnPasteText.Add_Click({
    $input = Show-PasteWindow
    if(-not [string]::IsNullOrWhiteSpace($input)){ Set-AutofillData (Parse-RequestText $input) "Pasted Text" }
})

$btnClearAuto.Add_Click({
    $Global:ParsedData = @{}; $txtAutoStatus.Text = "No Source Loaded (Manual Entry)"; $txtAutoStatus.Foreground = "#888"; $txtParsedPreview.Text = ""; $txtParsedPreview.Visibility = "Collapsed"; Log-Gui "Autofill cleared."
})

# --- TEMPLATE LOADER ---
function Load-Templates-List {
    Log-Gui "Scanning templates..."
    $jsons = Get-ChildItem -Path $TplDir -Filter *.json -File | Sort-Object Name
    $list=@()
    foreach($j in $jsons){
        try{
            $meta = Read-Json $j.FullName
            $msg=$null;$html=$null
            if($meta.PSObject.Properties['MsgFile']){$msg=Join-Path $TplDir $meta.MsgFile}
            if($meta.PSObject.Properties['HtmlFile']){$html=Join-Path $TplDir $meta.HtmlFile}
            if(-not $msg -and -not $html){
                $base=[IO.Path]::GetFileNameWithoutExtension($j.Name)
                $gMsg=Join-Path $TplDir ($base+'.msg'); $gHtml=Join-Path $TplDir ($base+'.html')
                if(Test-Path $gMsg){$msg=$gMsg}elseif(Test-Path $gHtml){$html=$gHtml}
            }
            if($msg -or $html){ $list += [pscustomobject]@{ Name = if($meta.Name){$meta.Name}else{$j.BaseName}; Group= if($meta.Group){$meta.Group}else{'General'}; Data = $meta; MsgPath=$msg; HtmlPath=$html } }
        } catch {}
    }
    $lstTemplates.ItemsSource = ($list | Sort-Object Group, Name)
}

function Run-Template {
    if($lstTemplates.SelectedItems.Count -eq 0){ Log-Gui "Select a template first."; return }
    $t = $lstTemplates.SelectedItems[0]
    Log-Gui "Preparing template: $($t.Name)"
    
    $dListEmails = Show-DListSelector
    if($dListEmails){ Log-Gui "Distribution Lists selected." }
    
    $vars = @{}
    function Smart-Get($Key, $Label){
        if($Global:ParsedData.ContainsKey($Key)){ return $Global:ParsedData[$Key] }
        if($Key -eq 'DateRange' -and $Global:ParsedData.ContainsKey('DateLine')){ return $Global:ParsedData['DateLine'] }
        if($Key -eq 'Title' -and $Global:ParsedData.ContainsKey('Subject')){ return $Global:ParsedData['Subject'] }
        return Show-PromptWindow "Template Input" "Enter value for: $Label" ""
    }

    if($t.Data.Requires){
        foreach($req in $t.Data.Requires){
            $val = Smart-Get $req $req
            if([string]::IsNullOrWhiteSpace($val) -and $val -ne "") { Log-Gui "Cancelled."; return }
            $vars[$req] = $val
        }
    }
    if(!$vars['DateRange']){ $vars['DateRange'] = Smart-Get "DateRange" "Date/Time Range" }
    if(!$vars['Title']){     $vars['Title']     = Smart-Get "Title" "Title/Subject" }
    foreach($k in $Global:ParsedData.Keys){ if(!$vars.ContainsKey($k)){ $vars[$k] = $Global:ParsedData[$k] } }

    $jobData = @{
        To = Safe-Join $t.Data.To; Cc = Safe-Join $t.Data.Cc
        Bcc = if((Safe-Join $t.Data.Bcc)){ "$(Safe-Join $t.Data.Bcc);$dListEmails" } else { $dListEmails }
        Subject = $t.Data.Subject; TemplateMsg = ""; Body = ""; BodyIsHtml = $false
    }
    if($t.MsgPath -and (Test-Path $t.MsgPath)){ $jobData.TemplateMsg = $t.MsgPath }
    elseif($t.HtmlPath -and (Test-Path $t.HtmlPath)){ $jobData.Body = Get-Content -Raw $t.HtmlPath; $jobData.BodyIsHtml = $true }
    else { $jobData.Body = "Error: Template body not found."; $jobData.BodyIsHtml = $false }

    Start-OutlookWorker -Vars $vars -JobData $jobData
}

# --- PATCHING & TOOLS ---
$Global:PatchData = @()
function Load-Patches {
    Log-Gui "Loading patches..."
    if (-not (Test-Path $PatchesExcel)) { Log-Gui "ERROR: Patches.xlsx missing!"; return }
    $lstPatches.ItemsSource = $null; $Global:PatchData = @()
    $xl = New-Object -ComObject Excel.Application; $xl.Visible = $false
    try {
        $wb = $xl.Workbooks.Open($PatchesExcel, $false, $true)
        $ws = $wb.Worksheets.Item("Patches")
        $rows = $ws.UsedRange.Rows.Count
        for($r=2; $r -le $rows; $r++){
            $rawDate = $ws.Cells.Item($r,4).Text
            $p = [pscustomobject]@{
                PatchId = (Coalesce $ws.Cells.Item($r,1).Text).Trim(); System = (Coalesce $ws.Cells.Item($r,2).Text).Trim(); Env = (Coalesce $ws.Cells.Item($r,3).Text).Trim()
                RawDate = $rawDate; DisplayDate = Fmt-Date $rawDate; DisplayTime = Fmt-Time $ws.Cells.Item($r,5).Text
                Servers = (Coalesce $ws.Cells.Item($r,6).Text).Trim(); To = (Coalesce $ws.Cells.Item($r,7).Text).Trim(); Cc = (Coalesce $ws.Cells.Item($r,8).Text).Trim()
                Status = if((Coalesce $ws.Cells.Item($r,2).Text).Trim()){"Ready"}else{"Incomplete"}
            }
            if($p.System){ $Global:PatchData += $p }
        }
        $Global:PatchData = $Global:PatchData | Sort-Object { [datetime]$_.RawDate }
        $lstPatches.ItemsSource = $Global:PatchData
        Log-Gui "Loaded $($Global:PatchData.Count) patch windows."
    } catch { Log-Gui "Error: $($_.Exception.Message)" } finally { if($wb){$wb.Close($false)}; $xl.Quit(); [GC]::Collect() }
}
function Generate-PatchEmail {
    param([string]$Type)
    if($lstPatches.SelectedItems.Count -eq 0){ Log-Gui "Select a row first."; return }
    $sel = $lstPatches.SelectedItems[0]
    $envTxt = if($sel.Env -match 'TEST|DEV|QA'){"- $($sel.Env) Servers"} elseif($sel.Env){"$($sel.Env) Servers"} else {"Servers"}
    $srvList = ($sel.Servers -split "[,;]+" | %{$_.Trim()} | ?{$_}) -join "`r`n"
    if ($Type -eq "BEFORE") {
        $subj = "Monthly Patching of $($sel.System) $envTxt - $($sel.DisplayDate) @ $($sel.DisplayTime)"
        $body = "Patching of $($sel.System) $envTxt will begin at $($sel.DisplayTime):`r`n$srvList"
    } else {
        $subj = "COMPLETE - Monthly Patching of $($sel.System) $envTxt - $($sel.DisplayDate)"
        $body = "Monthly patching and server reboots complete:`r`n" + $srvList
    }
    $job = @{ To=$sel.To; Cc=$sel.Cc; Bcc=""; Subject=$subj; Body=$body; BodyIsHtml=$false }
    Start-OutlookWorker -Vars @{} -JobData $job
}

function Run-CleanRDP { $dl = Join-Path $env:USERPROFILE "Downloads"; $c=0; Get-ChildItem $dl -Filter *.rdp | %{ Remove-Item $_.FullName -Force; $c++ }; Log-Gui "Cleaned $c files." }

function Run-PrintSchedules {
    # --- PART 1: DYNAMIC PATH DETECTION ---
    if (-not (Test-Path $SchedulesDir)) {
        $OneDrive = if ($env:OneDriveCommercial) { $env:OneDriveCommercial } else { $env:OneDrive }
        # SANITIZED: Replaced specific organization name with a generic placeholder
        $UserPath = "$env:USERPROFILE\OneDrive - [Organization]\Data Center Operations - Scripts\Email Suite\Data Center Operations - Operations Schedules"
        $GenericPath = "$OneDrive\Data Center Operations - Operations Schedules"

        if (Test-Path $UserPath) { $script:SchedulesDir = $UserPath }
        elseif ($OneDrive -and (Test-Path $GenericPath)) { $script:SchedulesDir = $GenericPath }
        else { 
            Log-Gui "Error: Schedules folder not found."
            [System.Windows.Forms.MessageBox]::Show("Could not find 'Data Center Operations - Operations Schedules'.`n`nWe looked here:`n$UserPath`n`nPlease ensure this folder is synced to your OneDrive.", "Folder Missing")
            return 
        }
    }

    # --- PART 2: COMPLEX PRINTING LOGIC ---
    Log-Gui "Printing from: $SchedulesDir"
    # SANITIZED: Replaced specific internal filename
    $TwoPageFile = "[Special_Print_Handling_Doc].xlsx"
    $files = Get-ChildItem $SchedulesDir -Include *.docx,*.xlsx -Recurse | ?{ $_.Attributes -notmatch "Directory" }
    
    if ($files.Count -eq 0) { Log-Gui "No files found."; return }

    $wordApp = $null
    if ($files.Name -match ".docx") {
        Log-Gui "Initializing Word application engine..."
        try { $wordApp = New-Object -ComObject Word.Application; $wordApp.Visible = $false } catch { Log-Gui "Error starting Word: $_" }
    }

    foreach($f in $files){ 
        Log-Gui "Processing $($f.Name)..."
        if ($f.Extension -eq ".docx") {
            try {
                if ($wordApp) { $doc = $wordApp.Documents.Open($f.FullName); $doc.PrintOut($false); $doc.Close($false) } 
                else { Start-Process -FilePath $f.FullName -Verb Print; Start-Sleep -Seconds 10 }
            } catch { Log-Gui "Failed to print $($f.Name): $_" }
        }
        elseif ($f.Name -match $TwoPageFile) {
            Log-Gui ">> ACTION: Select 'Properties' > '2-Sided' > OK"
            try { $excel = New-Object -ComObject Excel.Application; $excel.Visible = $true; $wb = $excel.Workbooks.Open($f.FullName); $excel.Dialogs.Item(8).Show(); $wb.Close($false); $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { Log-Gui "Error handling 2-sided print: $_" }
        } 
        else { Start-Process -FilePath $f.FullName -Verb Print; Start-Sleep -Seconds 5 }
    }

    if ($wordApp) { $wordApp.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wordApp) | Out-Null }
    Log-Gui "All print jobs sent."
}

# --- THE NEW SMART UPDATE FUNCTION (StartDT/EndDT) ---
# --- THE NEW SMART UPDATE FUNCTION (StartDT/EndDT) ---
function Run-UpdateDates {
    $targetYear = $txtYear.Text
    $targetMonth = $txtMonth.Text
    
    $FileChanges = (Resolve-Path $ChangesExcel).Path
    $FilePatches = (Resolve-Path $PatchesExcel).Path

    if (-not (Test-Path $FileChanges) -or -not (Test-Path $FilePatches)) { 
        Log-Gui "ERROR: Files not found."
        return 
    }

    Log-Gui "Starting Smart-Update for $targetMonth/$targetYear..."
    
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $true       
    $xl.DisplayAlerts = $false  
    $dateMap = @{} 
    $wb1 = $null
    $wb2 = $null

    try {
        # ---------------------------------------------------------
        # STEP 1: Process ChangeRequests.xlsx
        # ---------------------------------------------------------
        Log-Gui "Opening $FileChanges..."
        $wb1 = $xl.Workbooks.Open($FileChanges)
        $ws1 = $wb1.Worksheets.Item(1)
        
        # A. DYNAMIC COLUMN FINDER
        $colStart = 0; $colEnd = 0; $colSys = 0; $colFreq = 0; $colSrv = 0; $startRow = 0
        
        # Scan headers
        for ($r = 1; $r -le 10; $r++) {
            for ($c = 1; $c -le 20; $c++) {
                $txt = $ws1.Cells.Item($r, $c).Text
                if ($txt -match "StartDT")    { $colStart = $c }
                if ($txt -match "EndDT")      { $colEnd   = $c }
                if ($txt -match "System")     { $colSys   = $c }
                if ($txt -match "Server")     { $colSrv   = $c } 
                if ($txt -match "Frequency")  { $colFreq  = $c; $startRow = $r + 1 } 
            }
            if ($colFreq -gt 0) { break } 
        }

        if ($colFreq -eq 0 -or $colStart -eq 0) {
            Log-Gui "ERROR: Missing 'Frequency' or 'StartDT' headers in ChangeRequests."
            return
        }

        # B. PROCESS DATA
        $rows1 = $ws1.UsedRange.Rows.Count
        $updateCount = 0

        for ($r = $startRow; $r -le $rows1; $r++) {
            $sysName = $ws1.Cells.Item($r, $colSys).Text.Trim()
            $srvName = if ($colSrv -gt 0) { $ws1.Cells.Item($r, $colSrv).Text.Trim() } else { "" }
            $freqStr = $ws1.Cells.Item($r, $colFreq).Text
            
            $oldStartSerial = $ws1.Cells.Item($r, $colStart).Value2
            $oldEndSerial   = $ws1.Cells.Item($r, $colEnd).Value2

            if (($freqStr -match '(\d\w+|Last)\s+(\w+)\s+of the Month') -and ($oldStartSerial -is [double])) {
                $occurrence = $matches[1]; $dayName = $matches[2] 
                
                # Calculate the default date based on the text
                $newBaseDate = Get-MaintenanceDate -Year $targetYear -Month $targetMonth -Occurrence $occurrence -DayOfWeek $dayName
                
                # =========================================================
                # --- BUSINESS LOGIC OVERRIDE: CORE MED APP PROD vs TEST ---
                # =========================================================
                # SANITIZED: Replaced specific medical app name with generic placeholder
                if ($sysName -match "Core_Med_App - PROD" -and $occurrence -eq "3rd" -and $dayName -eq "Sun") {
                    # Figure out when TEST is patching (3rd Friday)
                    $testPatchDate = Get-MaintenanceDate -Year $targetYear -Month $targetMonth -Occurrence "3rd" -DayOfWeek "Fri"
                    
                    # If PROD's Sunday is BEFORE TEST's Friday...
                    if ($newBaseDate -lt $testPatchDate) {
                        Log-Gui "[$sysName] Date ($newBaseDate) falls before TEST ($testPatchDate). Shifting to 4th Sunday."
                        # Override to the 4th Sunday!
                        $newBaseDate = Get-MaintenanceDate -Year $targetYear -Month $targetMonth -Occurrence "4th" -DayOfWeek "Sun"
                    }
                }
                # =========================================================

                if ($newBaseDate) {
                    $timeOfDay = $oldStartSerial - [Math]::Floor($oldStartSerial)
                    $duration = if ($oldEndSerial -is [double]) { $oldEndSerial - $oldStartSerial } else { 0 }

                    $newBaseSerial  = $newBaseDate.ToOADate()
                    $newStartSerial = $newBaseSerial + $timeOfDay
                    $newEndSerial   = $newStartSerial + $duration

                    $ws1.Cells.Item($r, $colStart).Value2 = $newStartSerial
                    $ws1.Cells.Item($r, $colEnd).Value2   = $newEndSerial
                    
                    $cleanName = $sysName -replace " - ","-" 
                    $exactKey  = "$cleanName|$srvName"
                    
                    $dateMap[$exactKey] = $newBaseSerial
                    if (-not $dateMap.ContainsKey($cleanName)) { 
                        $dateMap[$cleanName] = $newBaseSerial 
                    }
                    $updateCount++
                }
            }
        }
        
        Log-Gui "Updates queued: $updateCount. Saving ChangeRequests..."
        $wb1.Save()
        
        # ---------------------------------------------------------
        # STEP 2: Process Patches.xlsx
        # ---------------------------------------------------------
        if ($updateCount -gt 0) {
            Log-Gui "Updating Patches.xlsx..."
            $wb2 = $xl.Workbooks.Open($FilePatches)
            $ws2 = $wb2.Worksheets.Item(1)
            $rows2 = $ws2.UsedRange.Rows.Count

            for ($r = 2; $r -le $rows2; $r++) {
                $pSys = $ws2.Cells.Item($r, 2).Text.Trim()
                $pEnv = $ws2.Cells.Item($r, 3).Text.Trim()
                $pSrv = $ws2.Cells.Item($r, 6).Text.Trim() 
                
                $lookupName = $pSys -replace " - ","-"
                $exactKey   = "$lookupName|$pSrv"
                
                $newVal = $null
                if ($dateMap.ContainsKey($exactKey)) {
                    $newVal = $dateMap[$exactKey]
                }
                elseif ($dateMap.ContainsKey($lookupName)) { 
                    $newVal = $dateMap[$lookupName] 
                }
                elseif ($pEnv -and $dateMap.ContainsKey("$lookupName-$pEnv")) { 
                    $newVal = $dateMap["$lookupName-$pEnv"] 
                }

                if ($newVal) {
                    $ws2.Cells.Item($r, 4).Value2 = $newVal
                }
            }
            $wb2.Save()
            Log-Gui "Patches.xlsx Saved."
        }
        
        [System.Windows.MessageBox]::Show("Smart Update Complete!`n`nUpdated $updateCount rows.", "Success")
    } 
    catch { 
        Log-Gui "FATAL ERROR: $($_.Exception.Message)" 
        [System.Windows.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
    } 
    finally { 
        if ($wb1) { $wb1.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb1)|Out-Null }
        if ($wb2) { $wb2.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb2)|Out-Null }
        if ($xl)  { $xl.Quit();  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        
        Load-Patches
    }
}

# --- HELPER FUNCTION (Ensures 3rd Wed, 2nd Sun logic works) ---
function Get-MaintenanceDate {
    param($Year, $Month, $Occurrence, $DayOfWeek)
    $targetDay = [DayOfWeek]$DayOfWeek
    $dt = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0
    while ($dt.DayOfWeek -ne $targetDay) { $dt = $dt.AddDays(1) }
    
    if ($Occurrence -match "1st") { $weeks = 0 }
    elseif ($Occurrence -match "2nd") { $weeks = 1 }
    elseif ($Occurrence -match "3rd") { $weeks = 2 }
    elseif ($Occurrence -match "4th") { $weeks = 3 }
    elseif ($Occurrence -match "Last") { 
        $weeks = 3 
        if ($dt.AddDays(28).Month -eq $Month) { $weeks = 4 }
    }
    return $dt.AddDays($weeks * 7)
}

# Events
$btnLoadPatches.Add_Click({ Load-Patches })
$btnPatchBefore.Add_Click({ Generate-PatchEmail "BEFORE" })
$btnPatchAfter.Add_Click({ Generate-PatchEmail "AFTER" })
$btnRunTemplate.Add_Click({ Run-Template })
$btnCleanRDP.Add_Click({ Run-CleanRDP })
$btnPrintSchedules.Add_Click({ Run-PrintSchedules })
$btnUpdateDates.Add_Click({ Run-UpdateDates })

# Init - Wrapped in CRASH GUARD
try {
    Load-Templates-List
    Log-Gui "System Ready. Welcome, $env:USERNAME."
    $window.ShowDialog() | Out-Null
} catch {
    Write-Host "FATAL STARTUP ERROR: $($_.Exception.Message)" -ForegroundColor Red
    [System.Windows.MessageBox]::Show("The application crashed on startup.`n`nError: $($_.Exception.Message)", "Startup Crash")
    Read-Host "Press Enter to exit..."
}