param(
    [switch]$AutoStart,
    [switch]$TestExpand
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ---------------------------------------------------------------- single instance guard
# Without this, every restart leaves a ghost tray icon behind and a second copy
# can be launched by both the scheduled task and a manual double-click.
$script:OrbMutex = New-Object System.Threading.Mutex($false, 'Global\Beacon_SingleInstance')
if (-not $script:OrbMutex.WaitOne(0)) {
    exit
}

# make the process DPI-aware so mouse hit-testing is correct on monitors
# whose scale factor differs from the primary display
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DpiAware {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
"@ -ErrorAction SilentlyContinue
try { [void][DpiAware]::SetProcessDPIAware() } catch {}

# ---------------------------------------------------------------- paths
$Base         = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Base) { $Base = 'D:\DeepSeek\workspace\orb' }
$SchedulePath = 'D:\DeepSeek\workspace\desktop-calendar\schedule.json'
$NotesPath    = Join-Path $Base 'notes.json'
$StatePath    = Join-Path $Base 'state.json'

# ---------------------------------------------------------------- helpers
function Brush([string]$hex) {
    New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
}

$COLORS = @{
    'Big Data Analytics'                                = '#FF9F0A'
    'Data Analysis'                                     = '#30D158'
    'Data Mining and AI'                                = '#BF5AF2'
    'Introduction to Data Science'                      = '#0A84FF'
    'Data Visualisation'                                = '#FF375F'
    'Research Methods and Dissertation Preparation'     = '#64D2FF'
    'Business Intelligence'                             = '#FFD60A'
    'Data and Society'                                  = '#AC8E68'
    'Data Science'                                      = '#5E5CE6'
    'Database Design'                                   = '#32ADE6'
}
$FALLBACK = @('#FF9F0A','#30D158','#BF5AF2','#0A84FF','#FF375F','#64D2FF','#FFD60A','#AC8E68','#5E5CE6','#32ADE6')

function ColorFor([string]$title) {
    if ($COLORS.ContainsKey($title)) { return $COLORS[$title] }
    $h = 0
    foreach ($ch in $title.ToCharArray()) { $h = (($h * 31) + [int]$ch) % 2147483647 }
    return $FALLBACK[$h % $FALLBACK.Length]
}

function ShortName([string]$t) {
    if ($COLORS.ContainsKey($t)) { return $t }
    if ($t.Contains(';')) { return '学院简报' }
    return $t
}

$SHORT = @{}

function ShortLabel([string]$t) {
    if (-not $t) { return '' }
    if ($t.Contains(';')) { return '简报' }

    $words = @($t -split '\s+' | Where-Object { $_ })
    if ($words.Count -le 3) { return $t }

    $take = @($words[0], $words[1], $words[2])
    # drop trailing connectors so we do not end on "and" / "to" / "of"
    while ($take.Count -gt 1 -and $take[-1] -match '^(and|to|of|the|a|an|in|for)$') {
        $take = @($take[0..($take.Count - 2)])
    }
    return ($take -join ' ')
}

# ---------------------------------------------------------------- data
$Schedule = @()
if (Test-Path $SchedulePath) {
    try {
        $rawJson = Get-Content $SchedulePath -Raw -Encoding UTF8
        $parsed = $rawJson | ConvertFrom-Json
        $Schedule = @($parsed)          # variable form expands; pipeline form would nest
        # PS 5.1 ConvertFrom-Json silently turns ISO date strings into DateTime
        foreach ($e in $Schedule) {
            $v = $e.date
            if ($v -isnot [string]) { $e.date = ([datetime]$v).ToString('yyyy-MM-dd') }
            $v = $e.start
            if ($v -isnot [string]) { $e.start = ([datetime]$v).ToString('yyyy-MM-ddTHH:mm:ss') }
            $v = $e.end
            if ($null -ne $v -and $v -isnot [string]) { $e.end = ([datetime]$v).ToString('yyyy-MM-ddTHH:mm:ss') }
        }
    } catch { $Schedule = @() }
}

# ---------------------------------------------------------------- imported ics sources
$SourcesPath = Join-Path $Base 'imported.json'

function Load-Sources {
    if (Test-Path $SourcesPath) {
        try {
            $raw = Get-Content $SourcesPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim()) {
                $o = $raw | ConvertFrom-Json
                if ($o -and $o.sources) { return @($o.sources) }
            }
        } catch {}
    }
    return @()
}

function Save-Sources($items) {
    $arr = @()
    if ($null -ne $items) { $arr = @($items | Where-Object { $null -ne $_ }) }
    $obj = [pscustomobject]@{ sources = $arr }
    $json = $obj | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($SourcesPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

$script:Sources = Load-Sources

function Rebuild-Events {
    $all = @($Schedule)
    foreach ($s in $script:Sources) {
        if ($s.events) { $all += @($s.events) }
    }
    $script:Events = $all
}
Rebuild-Events

function Unescape-Ics([string]$s) {
    if (-not $s) { return '' }
    return $s.Replace('\n', ' ').Replace('\N', ' ').Replace('\,', ',').Replace('\;', ';').Replace('\\', '\')
}

function Convert-IcsDate([string]$s) {
    if (-not $s) { return $null }
    $s = $s.Trim().TrimEnd('Z')
    if ($s.Length -ge 15) {
        try { return [datetime]::ParseExact($s.Substring(0,15), 'yyyyMMddTHHmmss', $null) } catch {}
    }
    if ($s.Length -ge 8) {
        try { return [datetime]::ParseExact($s.Substring(0,8), 'yyyyMMdd', $null) } catch {}
    }
    return $null
}

function Convert-VEvent($e) {
    $dtstart = $null
    $dtend   = $null
    foreach ($k in $e.Keys) {
        if ($k -like 'DTSTART*') { $dtstart = $e[$k] }
        elseif ($k -like 'DTEND*') { $dtend = $e[$k] }
    }
    if (-not $dtstart) { return $null }

    $st = Convert-IcsDate $dtstart
    if (-not $st -or $st.Year -lt 2000) { return $null }
    $en = $null
    if ($dtend) { $en = Convert-IcsDate $dtend }

    $sum = ''
    if ($e.ContainsKey('SUMMARY')) { $sum = $e['SUMMARY'] }
    $loc = ''
    if ($e.ContainsKey('LOCATION')) { $loc = $e['LOCATION'] }

    $sum = ((Unescape-Ics $sum).Trim() -replace '\s+', ' ')
    $loc = ((Unescape-Ics $loc).Trim() -replace '\s+', ' ')

    return [pscustomobject]@{
        title    = $sum
        start    = $st.ToString('yyyy-MM-ddTHH:mm:ss')
        end      = $(if ($en) { $en.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null })
        location = $loc
        date     = $st.ToString('yyyy-MM-dd')
        time     = $st.ToString('HH:mm')
    }
}

function Parse-IcsText([string]$text) {
    $lines = $text -split "`r?`n"
    $un = New-Object System.Collections.ArrayList
    foreach ($l in $lines) {
        if ($l -match '^[ \t]' -and $un.Count -gt 0) {
            $un[$un.Count - 1] = $un[$un.Count - 1] + $l.Substring(1)
        } else { [void]$un.Add($l) }
    }

    $res = @()
    $cur = $null
    foreach ($l in $un) {
        if ($l -eq 'BEGIN:VEVENT') { $cur = @{} }
        elseif ($l -eq 'END:VEVENT') {
            if ($null -ne $cur) {
                $ev = Convert-VEvent $cur
                if ($ev) { $res += $ev }
            }
            $cur = $null
        }
        elseif ($null -ne $cur -and $l.Contains(':')) {
            $i = $l.IndexOf(':')
            $k = $l.Substring(0, $i)
            $v = $l.Substring($i + 1)
            if (-not $cur.ContainsKey($k)) { $cur[$k] = $v }
        }
    }
    return $res
}

function Load-Notes {
    if (Test-Path $NotesPath) {
        try {
            $raw = Get-Content $NotesPath -Raw -Encoding UTF8
            if ($raw -and $raw.Trim()) {
                $o = $raw | ConvertFrom-Json
                if ($o -and $o.items) {
                    $items = @($o.items)
                    foreach ($it in $items) {
                        # ConvertFrom-Json may have turned the due string into DateTime
                        if ($null -ne $it.due -and $it.due -isnot [string]) {
                            $it.due = ([datetime]$it.due).ToString('yyyy-MM-dd')
                        }
                    }
                    return $items
                }
            }
        } catch {}
    }
    return @()
}

function Save-Notes($items) {
    $arr = @()
    if ($null -ne $items) { $arr = @($items | Where-Object { $null -ne $_ }) }
    $obj = [pscustomobject]@{ items = $arr }
    $json = $obj | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($NotesPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

$Notes = Load-Notes

# pending due date for the note being written; null = no date
$script:DueDate  = $null
$script:CalYear  = (Get-Date).Year
$script:CalMonth = (Get-Date).Month

function Format-DueShort($key) {
    if (-not $key) { return '' }
    try {
        $dt = [datetime]::ParseExact([string]$key, 'yyyy-MM-dd', $null)
        return "$($MON[$dt.Month-1])$($dt.Day)日"
    } catch { return [string]$key }
}

function Load-State {
    $s = @{ Left = $null; Top = $null }
    if (Test-Path $StatePath) {
        try {
            $o = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $o.Left) { $s.Left = [double]$o.Left }
            if ($null -ne $o.Top)  { $s.Top  = [double]$o.Top }
        } catch {}
    }
    return $s
}

function Save-State {
    $o = [pscustomobject]@{ Left = [double]$w.Left; Top = [double]$w.Top }
    [System.IO.File]::WriteAllText($StatePath, ($o | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
}

# ---------------------------------------------------------------- xaml
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Orb" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False" ResizeMode="NoResize"
        Width="470" Height="76"
        FontFamily="Microsoft YaHei UI, Segoe UI, sans-serif">
  <Window.Resources>
    <Style x:Key="TabBtn" TargetType="Button">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="11,5">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="{x:Type ScrollBar}">
      <Setter Property="Width" Value="8"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type ScrollBar}">
            <Border Background="Transparent" Width="8">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="{x:Type Thumb}">
                        <Border CornerRadius="4" Background="#38FFFFFF" Margin="1,2"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>
  <Grid x:Name="Root">

    <Border x:Name="OrbView" Width="470" Height="76" CornerRadius="14"
            Background="#F0181B21" BorderBrush="#42FFFFFF" BorderThickness="1">
      <Grid Margin="15,0,10,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="72"/>
          <ColumnDefinition Width="184"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="48"/>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" VerticalAlignment="Center" Margin="0,0,2,0">
          <TextBlock x:Name="ClockText" Text="--:--" Foreground="#FFFFFF" FontSize="17"
                     FontWeight="SemiBold" HorizontalAlignment="Left"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" Margin="0,2,0,0">
            <TextBlock x:Name="OrbDay" Text="15" Foreground="#9A9AA0" FontSize="11"/>
            <TextBlock Text="·" Foreground="#4E4E54" FontSize="11" Margin="3,0"/>
            <TextBlock x:Name="OrbWd" Text="周日" Foreground="#6E6E73" FontSize="11"/>
          </StackPanel>
        </StackPanel>

        <Grid Grid.Column="1" VerticalAlignment="Center" Margin="0,0,16,0">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,4">
            <TextBlock x:Name="CpuText" Text="CPU --" Width="52" Foreground="#D8D8DC" FontSize="11"/>
            <TextBlock x:Name="TempText" Text="--°C" Width="42" Margin="4,0,0,0" Foreground="#6E6E73" FontSize="11"/>
            <TextBlock x:Name="GpuText" Text="GPU --" Width="56" Margin="8,0,0,0" Foreground="#D8D8DC" FontSize="11"/>
          </StackPanel>

          <StackPanel Grid.Row="1" Orientation="Horizontal">
            <TextBlock x:Name="RamText" Text="RAM --" Width="98" Foreground="#8A8A90" FontSize="11"/>
            <TextBlock x:Name="WeatherText" Text="--" Width="72" Margin="4,0,0,0" Foreground="#8A8A90" FontSize="11"/>
          </StackPanel>
        </Grid>

        <Grid Grid.Column="2" VerticalAlignment="Center" Margin="0,0,10,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Ellipse x:Name="DueDot" Grid.Column="0" Width="7" Height="7" Fill="#30D158"
                   Visibility="Collapsed" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <StackPanel Grid.Column="1">
            <TextBlock x:Name="NextMain" Text="—" Foreground="#F5F5F7" FontSize="11.5"
                       FontWeight="Medium" TextTrimming="CharacterEllipsis"/>
            <TextBlock x:Name="NextSub" Text="" Foreground="#6E6E73" FontSize="9.5"
                       Margin="0,3,0,0" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
        </Grid>

        <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
          <Border x:Name="LockBtn" Width="24" Height="24" Background="Transparent"
                  CornerRadius="5" Cursor="Hand">
            <TextBlock x:Name="LockIcon" Text="&#xE785;" FontFamily="Segoe MDL2 Assets"
                       FontSize="12" Foreground="#7A7A80"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <Border x:Name="LayerBtn" Width="24" Height="24" Background="Transparent"
                  CornerRadius="5" Cursor="Hand" Margin="2,0,0,0">
            <TextBlock x:Name="LayerIcon" Text="&#xE718;" FontFamily="Segoe MDL2 Assets"
                       FontSize="12" Foreground="#0A84FF"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
        </StackPanel>
      </Grid>
    </Border>

    <Border x:Name="PanelView" Visibility="Collapsed" Width="380" Height="560"
            CornerRadius="16" Background="#F5141619" BorderBrush="#2EFFFFFF" BorderThickness="1">
      <Grid x:Name="PanelGrid">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="190"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="20,16,14,10">
          <StackPanel VerticalAlignment="Center">
            <TextBlock x:Name="HeadDate" Text="" Foreground="#F5F5F7" FontSize="16.5" FontWeight="SemiBold"/>
            <TextBlock x:Name="HeadSub" Text="" Foreground="#8A8A90" FontSize="11.5" Margin="0,3,0,0"/>
          </StackPanel>

          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,36,0">
            <Button x:Name="TabMonth" Content="月历" Style="{StaticResource TabBtn}"
                    Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                    FontSize="10.5" Cursor="Hand"/>
            <Button x:Name="TabSchedule" Content="日程" Style="{StaticResource TabBtn}" Margin="4,0,0,0"
                    Background="#33FFFFFF" Foreground="#F5F5F7" BorderThickness="0"
                    FontSize="10.5" Cursor="Hand"/>
            <Button x:Name="SettingsBtn" Style="{StaticResource TabBtn}" Margin="10,0,0,0"
                    Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                    Cursor="Hand" Padding="8,0">
              <TextBlock Text="&#xE713;" FontFamily="Segoe MDL2 Assets" FontSize="12"/>
            </Button>
          </StackPanel>

          <Button x:Name="CloseBtn" Content="✕" HorizontalAlignment="Right" VerticalAlignment="Center"
                  Width="28" Height="28" Background="Transparent" Foreground="#8A8A90"
                  BorderThickness="0" FontSize="13" Cursor="Hand" Padding="0"/>
        </Grid>

        <Grid Grid.Row="1">
          <ScrollViewer x:Name="ScheduleView" VerticalScrollBarVisibility="Auto" Margin="16,0,10,0">
            <StackPanel x:Name="ClassList" Margin="0,0,6,0"/>
          </ScrollViewer>

          <Grid x:Name="MonthView" Visibility="Collapsed" Margin="16,2,12,6">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Grid Grid.Row="0" Margin="0,0,0,8">
              <TextBlock x:Name="MvTitle" FontSize="12.5" FontWeight="SemiBold" Foreground="#F5F5F7"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
              <Button x:Name="MvPrev" Content="‹" Width="24" Height="22" HorizontalAlignment="Left"
                      Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                      FontSize="13" Cursor="Hand" Padding="0"/>
              <Button x:Name="MvNext" Content="›" Width="24" Height="22" HorizontalAlignment="Right"
                      Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                      FontSize="13" Cursor="Hand" Padding="0"/>
            </Grid>

            <UniformGrid Grid.Row="1" Columns="7" Margin="0,0,0,4">
              <TextBlock Text="日" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="一" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="二" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="三" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="四" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="五" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
              <TextBlock Text="六" Foreground="#5E5E64" FontSize="9.5" HorizontalAlignment="Center"/>
            </UniformGrid>

            <UniformGrid x:Name="MvGrid" Grid.Row="2" Columns="7" Rows="6" Height="228"/>

            <Grid Grid.Row="3" Margin="0,10,0,0">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <TextBlock x:Name="MvDayTitle" Grid.Row="0" Text="" Foreground="#8A8A90"
                         FontSize="11" Margin="2,0,0,6"/>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <StackPanel x:Name="MvDayEvents"/>
              </ScrollViewer>
            </Grid>
          </Grid>
        </Grid>

        <Grid x:Name="NoteHeaderRow" Grid.Row="2" Margin="20,14,18,6">
          <TextBlock Text="备忘" Foreground="#8A8A90" FontSize="11.5" VerticalAlignment="Center"/>
          <TextBlock x:Name="NoteCount" Text="" Foreground="#5E5E64" FontSize="11"
                     HorizontalAlignment="Right" VerticalAlignment="Center"/>
        </Grid>

        <Grid x:Name="NoteAreaRow" Grid.Row="3" Margin="16,0,16,16">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Border Grid.Row="0" CornerRadius="9" Background="#1FFFFFFF" Margin="0,0,0,6">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="NoteInput" Grid.Column="0" Background="Transparent" BorderThickness="0"
                       Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="11,9"
                       FontSize="12.5" VerticalContentAlignment="Center"/>
              <TextBlock x:Name="NoteHint" Grid.Column="0" Text="写点什么…" IsHitTestVisible="False"
                         Foreground="#5E5E64" FontSize="12.5" Margin="13,0,0,0"
                         VerticalAlignment="Center"/>
              <Button x:Name="DueBtn" Grid.Column="1" Content="&#x1F4C5;" Width="30" Height="30"
                      Margin="0,0,1,0" Background="Transparent" Foreground="#8A8A90"
                      BorderThickness="0" FontSize="13" Cursor="Hand"
                      VerticalAlignment="Center" Padding="0"/>
              <Button x:Name="AddBtn" Grid.Column="2" Content="+" Width="34" Height="30"
                      Margin="0,0,3,0" Background="Transparent" Foreground="#0A84FF"
                      BorderThickness="0" FontSize="18" FontWeight="SemiBold" Cursor="Hand"
                      VerticalAlignment="Center" Padding="0"/>
            </Grid>
          </Border>

          <Border x:Name="DueChip" Grid.Row="1" Visibility="Collapsed"
                  HorizontalAlignment="Left" CornerRadius="6" Background="#26FF9F0A"
                  Padding="8,3" Margin="2,0,0,6">
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="DueText" Foreground="#FF9F0A" FontSize="10.5" VerticalAlignment="Center"/>
              <TextBlock x:Name="DueClear" Text="✕" Foreground="#FF9F0A" FontSize="9.5"
                         Cursor="Hand" Margin="6,1,0,0" VerticalAlignment="Center"/>
            </StackPanel>
          </Border>

          <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="NoteList"/>
          </ScrollViewer>
        </Grid>
      </Grid>
    </Border>

    <Popup x:Name="CalPopup" StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade" Placement="Top">
      <Border CornerRadius="12" Background="#F51A1D23" BorderBrush="#3AFFFFFF" BorderThickness="1" Padding="12">
        <StackPanel>
          <Grid Margin="0,0,0,10">
            <TextBlock x:Name="CalTitle" Foreground="#F5F5F7" FontSize="13" FontWeight="SemiBold"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
            <Button x:Name="CalPrev" Content="‹" Width="26" Height="24" HorizontalAlignment="Left"
                    Background="Transparent" Foreground="#8A8A90" BorderThickness="0" FontSize="14" Cursor="Hand" Padding="0"/>
            <Button x:Name="CalNext" Content="›" Width="26" Height="24" HorizontalAlignment="Right"
                    Background="Transparent" Foreground="#8A8A90" BorderThickness="0" FontSize="14" Cursor="Hand" Padding="0"/>
          </Grid>
          <UniformGrid x:Name="CalGrid" Columns="7" Rows="7"/>
        </StackPanel>
      </Border>
    </Popup>
  </Grid>
</Window>
'@

# ---------------------------------------------------------------- settings window xaml
$settingsXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="日程球 设置" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False" ResizeMode="NoResize"
        Width="420" Height="430"
        FontFamily="Microsoft YaHei UI, Segoe UI, sans-serif">
  <Window.Resources>
    <Style TargetType="{x:Type ScrollBar}">
      <Setter Property="Width" Value="8"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type ScrollBar}">
            <Border Background="Transparent" Width="8">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb>
                  <Thumb>
                    <Thumb.Template>
                      <ControlTemplate TargetType="{x:Type Thumb}">
                        <Border CornerRadius="4" Background="#38FFFFFF" Margin="1,2"/>
                      </ControlTemplate>
                    </Thumb.Template>
                  </Thumb>
                </Track.Thumb>
              </Track>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Border CornerRadius="16" Background="#FC15181E" BorderBrush="#2EFFFFFF" BorderThickness="1">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <Grid x:Name="SetHeader" Grid.Row="0" Margin="20,16,14,12" Background="Transparent">
        <TextBlock Text="设置" Foreground="#F5F5F7" FontSize="15" FontWeight="SemiBold"
                   VerticalAlignment="Center"/>
        <Button x:Name="SetCloseBtn" Content="✕" HorizontalAlignment="Right" Width="28" Height="28"
                Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                FontSize="13" Cursor="Hand" Padding="0"/>
      </Grid>

      <StackPanel Grid.Row="1" Margin="20,0,20,14">
        <TextBlock Text="导入日程链接（.ics）" Foreground="#8A8A90" FontSize="11" Margin="0,0,0,8"/>
        <Border CornerRadius="9" Background="#1FFFFFFF">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="SetUrlInput" Grid.Column="0" Background="Transparent" BorderThickness="0"
                     Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="11,9" FontSize="11.5"/>
            <TextBlock x:Name="SetUrlHint" Grid.Column="0" Text="https://…" IsHitTestVisible="False"
                       Foreground="#4E4E54" FontSize="11.5" Margin="13,0,0,0" VerticalAlignment="Center"/>
            <Button x:Name="SetImportBtn" Grid.Column="1" Content="导入" Width="54" Height="30"
                    Margin="0,0,4,0" Background="Transparent" Foreground="#0A84FF"
                    BorderThickness="0" FontSize="11.5" Cursor="Hand" Padding="0"
                    VerticalAlignment="Center"/>
          </Grid>
        </Border>
        <TextBlock x:Name="SetStatus" Text="" Foreground="#8A8A90" FontSize="10.5"
                   Margin="2,8,0,0" TextWrapping="Wrap"/>
      </StackPanel>

      <TextBlock Grid.Row="2" Text="已导入的日历" Foreground="#8A8A90" FontSize="11"
                 Margin="20,0,0,8"/>

      <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="20,0,16,16">
        <StackPanel x:Name="SetSourceList"/>
      </ScrollViewer>
    </Grid>
  </Border>
</Window>
'@

$w = [System.Windows.Markup.XamlReader]::Parse($xaml)

# window icon (taskbar / alt-tab)
try {
    $winIcon = Join-Path $Base 'assets\beacon.ico'
    if (-not (Test-Path $winIcon)) { $winIcon = Join-Path $Base 'app\assets\beacon.ico' }
    if (Test-Path $winIcon) {
        $w.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([uri]$winIcon)
    }
} catch {}

$OrbView   = $w.FindName('OrbView')
$PanelView = $w.FindName('PanelView')
$PanelGrid = $w.FindName('PanelGrid')
$OrbDay    = $w.FindName('OrbDay')
$OrbWd     = $w.FindName('OrbWd')
$ClockText = $w.FindName('ClockText')
$CpuText   = $w.FindName('CpuText')
$TempText  = $w.FindName('TempText')
$GpuText   = $w.FindName('GpuText')
$RamText   = $w.FindName('RamText')
$WeatherText = $w.FindName('WeatherText')
$NextMain  = $w.FindName('NextMain')
$NextSub   = $w.FindName('NextSub')
$DueDot    = $w.FindName('DueDot')
$ScheduleView = $w.FindName('ScheduleView')
$MonthView    = $w.FindName('MonthView')
$TabSchedule  = $w.FindName('TabSchedule')
$TabMonth     = $w.FindName('TabMonth')
$MvTitle      = $w.FindName('MvTitle')
$MvGrid       = $w.FindName('MvGrid')
$MvDayTitle   = $w.FindName('MvDayTitle')
$MvDayEvents  = $w.FindName('MvDayEvents')
$MvPrev       = $w.FindName('MvPrev')
$MvNext       = $w.FindName('MvNext')
$SettingsBtn  = $w.FindName('SettingsBtn')
$LockBtn      = $w.FindName('LockBtn')
$LockIcon     = $w.FindName('LockIcon')
$LayerBtn     = $w.FindName('LayerBtn')
$LayerIcon    = $w.FindName('LayerIcon')
$NoteHeaderRow = $w.FindName('NoteHeaderRow')
$NoteAreaRow   = $w.FindName('NoteAreaRow')

# settings window
$sw = [System.Windows.Markup.XamlReader]::Parse($settingsXaml)
$SetHeader     = $sw.FindName('SetHeader')
$SetCloseBtn   = $sw.FindName('SetCloseBtn')
$SetUrlInput   = $sw.FindName('SetUrlInput')
$SetUrlHint    = $sw.FindName('SetUrlHint')
$SetImportBtn  = $sw.FindName('SetImportBtn')
$SetStatus     = $sw.FindName('SetStatus')
$SetSourceList = $sw.FindName('SetSourceList')
$HeadDate  = $w.FindName('HeadDate')
$HeadSub   = $w.FindName('HeadSub')
$CloseBtn  = $w.FindName('CloseBtn')
$ClassList = $w.FindName('ClassList')
$NoteList  = $w.FindName('NoteList')
$NoteInput = $w.FindName('NoteInput')
$NoteHint  = $w.FindName('NoteHint')
$NoteCount = $w.FindName('NoteCount')
$AddBtn    = $w.FindName('AddBtn')
$DueBtn    = $w.FindName('DueBtn')
$DueChip   = $w.FindName('DueChip')
$DueText   = $w.FindName('DueText')
$DueClear  = $w.FindName('DueClear')
$CalPopup  = $w.FindName('CalPopup')
$CalTitle  = $w.FindName('CalTitle')
$CalGrid   = $w.FindName('CalGrid')
$CalPrev   = $w.FindName('CalPrev')
$CalNext   = $w.FindName('CalNext')

# ---------------------------------------------------------------- placement
$wa = [System.Windows.SystemParameters]::WorkArea
$state = Load-State
if ($null -ne $state.Left -and $null -ne $state.Top) {
    $w.Left = $state.Left
    $w.Top  = $state.Top
} else {
    # Prefer the user's spot on the secondary monitor (-1912, 7).
    # If that point does not fall inside any connected display, fall back to
    # the top-right corner of the primary screen.
    $prefLeft = -1912
    $prefTop  = 7
    $prefOk   = $false
    foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
        $b = $sc.Bounds
        if ($prefLeft -ge ($b.Left - 4)  -and $prefLeft -le ($b.Right - 200) -and
            $prefTop  -ge ($b.Top  - 4)  -and $prefTop  -le ($b.Bottom - 100)) {
            $prefOk = $true
            break
        }
    }
    if ($prefOk) {
        $w.Left = $prefLeft
        $w.Top  = $prefTop
    } else {
        $w.Left = $wa.Right - 470 - 28
        $w.Top  = $wa.Top + 20
    }
}

# keep it on some real screen; otherwise fall back to the primary corner
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
$onScreen = $false
foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
    $b = $sc.Bounds
    if ($w.Left -ge ($b.Left - 40)  -and $w.Left -le ($b.Right - 120) -and
        $w.Top  -ge ($b.Top  - 10)  -and $w.Top  -le ($b.Bottom - 40)) { $onScreen = $true; break }
}
if (-not $onScreen) {
    $w.Left = $wa.Right - 470 - 28
    $w.Top  = $wa.Top + 150
}

# ---------------------------------------------------------------- metrics
$script:RamTotalGB = 31.1
try {
    $os0 = Get-CimInstance Win32_OperatingSystem
    $script:RamTotalGB = [math]::Round($os0.TotalVisibleMemorySize / 1MB, 1)
} catch {}

$script:pcCpu = $null
try {
    $script:pcCpu = New-Object System.Diagnostics.PerformanceCounter('Processor', '% Processor Time', '_Total')
    $null = $script:pcCpu.NextValue()
} catch {}

$script:pcMem = $null
try {
    $script:pcMem = New-Object System.Diagnostics.PerformanceCounter('Memory', 'Available MBytes')
    $null = $script:pcMem.NextValue()
} catch {}

# ---------------------------------------------------------------- cpu temperature via LibreHardwareMonitor
$script:LhmComputer = $null
try {
    $libDir = Join-Path $Base 'lib'
    foreach ($n in @(
        'System.Runtime.CompilerServices.Unsafe.dll',
        'System.Buffers.dll',
        'System.Numerics.Vectors.dll',
        'System.Threading.Tasks.Extensions.dll',
        'System.Memory.dll',
        'HidSharp.dll',
        'LibreHardwareMonitorLib.dll')) {
        $p = Join-Path $libDir $n
        if (Test-Path $p) {
            try { [void][System.Reflection.Assembly]::Load([System.IO.File]::ReadAllBytes($p)) } catch {}
        }
    }
    $lhm = New-Object LibreHardwareMonitor.Hardware.Computer
    $lhm.IsCpuEnabled = $true
    $lhm.IsGpuEnabled = $true
    $lhm.Open()
    $script:LhmComputer = $lhm
} catch {
    $script:LhmComputer = $null
}

function Get-LhmCpuTemp {
    if (-not $script:LhmComputer) { return $null }
    try {
        foreach ($h in $script:LhmComputer.Hardware) {
            if ("$($h.HardwareType)" -ne 'Cpu') { continue }
            $h.Update()
            $best = $null
            foreach ($s in $h.Sensors) {
                if ("$($s.SensorType)" -ne 'Temperature') { continue }
                if ($null -eq $s.Value -or $s.Value -le 1) { continue }
                if ($s.Name -like 'Core (Tctl/Tdie)*') { return [double]$s.Value }
                if ($s.Name -eq 'CPU Package' -and $null -eq $best) { $best = [double]$s.Value }
                if ($s.Name -like 'Core #1*' -and $null -eq $best) { $best = [double]$s.Value }
            }
            if ($null -ne $best) { return $best }
        }
    } catch {}
    return $null
}

function Get-LhmGpuLoad {
    if (-not $script:LhmComputer) { return $null }
    try {
        # Machines can have several adapters (e.g. AMD iGPU + NVIDIA dGPU, or a
        # laptop with only integrated graphics). Take the busiest "GPU Core" load
        # so we always show the adapter that is actually working.
        $max = $null
        foreach ($h in $script:LhmComputer.Hardware) {
            if ("$($h.HardwareType)" -notmatch 'Gpu') { continue }
            $h.Update()
            foreach ($s in $h.Sensors) {
                if ("$($s.SensorType)" -ne 'Load') { continue }
                if ("$($s.Name)" -ne 'GPU Core') { continue }
                if ($null -eq $s.Value) { continue }
                $v = [double]$s.Value
                if ($null -eq $max -or $v -gt $max) { $max = $v }
            }
        }
        return $max
    } catch {}
    return $null
}

# ---------------------------------------------------------------- weather
$script:WeatherCache = '--'
$script:WeatherAt    = [datetime]::MinValue

function Get-WeatherDesc([int]$code) {
    if ($code -eq 0) { return '晴' }
    if ($code -eq 1 -or $code -eq 2) { return '多云' }
    if ($code -eq 3) { return '阴' }
    if ($code -eq 45 -or $code -eq 48) { return '雾' }
    if ($code -ge 51 -and $code -le 57) { return '毛毛雨' }
    if ($code -ge 61 -and $code -le 67) { return '雨' }
    if ($code -ge 71 -and $code -le 77) { return '雪' }
    if ($code -ge 80 -and $code -le 82) { return '阵雨' }
    if ($code -eq 85 -or $code -eq 86) { return '阵雪' }
    if ($code -ge 95) { return '雷雨' }
    return ''
}

function Update-Weather {
    if (((Get-Date) - $script:WeatherAt).TotalMinutes -lt 20) { return }
    $script:WeatherAt = Get-Date
    try {
        $r = Invoke-RestMethod -Uri 'https://api.open-meteo.com/v1/forecast?latitude=53.3811&longitude=-1.4701&current=temperature_2m,weather_code' -TimeoutSec 8
        $t = [math]::Round([double]$r.current.temperature_2m)
        $desc = Get-WeatherDesc ([int]$r.current.weather_code)
        $script:WeatherCache = "$t°C $desc"
    } catch {
        if ($script:WeatherCache -eq '--') { $script:WeatherCache = '天气 --' }
    }
}

$script:NvSmi = $null -ne (Get-Command nvidia-smi -ErrorAction SilentlyContinue)

function Update-NextThing {
    $now = Get-Date
    $todayKey = $now.ToString('yyyy-MM-dd')
    $nowT = $now.ToString('HH:mm')
    $tomorrowKey = $now.AddDays(1).ToString('yyyy-MM-dd')

    $best = $null
    $bestKey = ''

    foreach ($e in $script:Events) {
        if ($e.date -lt $todayKey) { continue }
        if ($e.date -eq $todayKey -and $e.time -lt $nowT) { continue }
        $k = "$($e.date) $($e.time)"
        if ($null -eq $best -or $k -lt $bestKey) {
            $dt = [datetime]::ParseExact($e.date, 'yyyy-MM-dd', $null)
            $dayLabel = $(if ($e.date -eq $todayKey) { '今天' } elseif ($e.date -eq $tomorrowKey) { '明天' } else { "$($dt.Month)/$($dt.Day)" })
            $bestKey = $k
            $best = [pscustomobject]@{ Main = (ShortLabel $e.title); Sub = "$dayLabel $($e.time)"; Kind = 'class' }
        }
    }

    foreach ($n in $script:Notes) {
        if ($n.done -or -not $n.due) { continue }
        $d = [string]$n.due
        if ($d -lt $todayKey) { continue }
        $k = "$d 23:59"
        if ($null -eq $best -or $k -lt $bestKey) {
            $dd = [datetime]::ParseExact($d, 'yyyy-MM-dd', $null)
            $dayLabel = $(if ($d -eq $todayKey) { '今天' } elseif ($d -eq $tomorrowKey) { '明天' } else { "$($dd.Month)/$($dd.Day)" })
            $leftDays = ($dd.Date - $now.Date).Days
            $dotColor = $(if ($leftDays -le 0) { '#FF453A' } elseif ($leftDays -le 3) { '#FFD60A' } else { '#30D158' })
            $bestKey = $k
            $best = [pscustomobject]@{ Main = $n.text; Sub = "截止 $dayLabel"; Kind = 'due'; DotColor = $dotColor }
        }
    }

    if ($null -eq $best) {
        $NextMain.Text = '—'
        $NextMain.ToolTip = ''
        $NextMain.Foreground = Brush '#6E6E73'
        $NextSub.Text = '没有近期安排'
        $DueDot.Visibility = 'Collapsed'
        return
    }

    $NextMain.Text = $best.Main
    $NextMain.ToolTip = $best.Main
    $NextSub.Text = $best.Sub

    if ($best.Kind -eq 'due') {
        $NextMain.Foreground = Brush '#FF9F0A'
        $DueDot.Visibility = 'Visible'
        $DueDot.Fill = Brush $best.DotColor
    } else {
        $NextMain.Foreground = Brush '#F5F5F7'
        $DueDot.Visibility = 'Collapsed'
    }
}

function Update-Orb {
    $now = Get-Date
    $OrbDay.Text = $now.Day.ToString()
    $OrbWd.Text  = $WD[[int]$now.DayOfWeek]
    if ($ClockText) { $ClockText.Text = $now.ToString('HH:mm') }

    if ($script:pcCpu) {
        try { $CpuText.Text = 'CPU ' + [int][math]::Round($script:pcCpu.NextValue()) + '%' }
        catch { $CpuText.Text = 'CPU --' }
    }

    # CPU temperature: LibreHardwareMonitor first, AIDA64 registry as fallback
    $temp = Get-LhmCpuTemp
    if ($null -eq $temp) {
        try {
            $so = Get-ItemProperty 'HKCU:\Software\FinalWire\AIDA64\SensorValues' -ErrorAction Stop
            foreach ($cand in @('Value.TCPU', 'Value.TCPU1', 'TCPU1', 'TCPU')) {
                $vv = $so.$cand
                if ($null -ne $vv -and "$vv".Trim() -match '^\d+(\.\d+)?$') {
                    $temp = [double]$vv
                    break
                }
            }
        } catch {}
    }
    if ($null -ne $temp) {
        $TempText.Text = ([int][math]::Round($temp, 0)).ToString() + '°C'
        $TempText.Foreground = Brush $(if ($temp -ge 85) { '#FF453A' } elseif ($temp -ge 70) { '#FF9F0A' } else { '#8A8A90' })
    } else {
        $TempText.Text = '--°C'
        $TempText.Foreground = Brush '#6E6E73'
    }

    # GPU: LibreHardwareMonitor handles both discrete and integrated GPUs (works on
    # laptops with only iGPU). nvidia-smi is the fallback when LHM has nothing.
    $gpuPct = Get-LhmGpuLoad
    if ($null -eq $gpuPct -and $script:NvSmi) {
        try {
            $o = & nvidia-smi --query-gpu=utilization.gpu --format=csv,noheader,nounits 2>$null
            if ($o) { $gpuPct = [double]([string]$o).Trim() }
        } catch {}
    }
    if ($null -ne $gpuPct) {
        $GpuText.Text = 'GPU ' + ([math]::Round($gpuPct, 0)).ToString() + '%'
        $GpuText.Foreground = Brush $(if ($gpuPct -ge 90) { '#FF453A' } elseif ($gpuPct -ge 60) { '#FF9F0A' } else { '#D8D8DC' })
    } else {
        $GpuText.Text = 'GPU --'
        $GpuText.Foreground = Brush '#6E6E73'
    }

    if ($script:pcMem) {
        try {
            $availGB = $script:pcMem.NextValue() / 1024
            $usedGB  = $script:RamTotalGB - $availGB
            $RamText.Text = 'RAM ' + ([math]::Round($usedGB, 1)).ToString() + '/' +
                            $script:RamTotalGB.ToString()
        } catch { }
    }

    Update-Weather
    $WeatherText.Text = $script:WeatherCache

    Update-NextThing
}

# ---------------------------------------------------------------- month view
$script:MvYear  = (Get-Date).Year
$script:MvMonth = (Get-Date).Month

function Switch-Tab([string]$name) {
    if ($name -eq 'month') {
        $ScheduleView.Visibility = 'Collapsed'
        $MonthView.Visibility    = 'Visible'
        $TabMonth.Background     = Brush '#33FFFFFF'
        $TabMonth.Foreground     = Brush '#F5F5F7'
        $TabSchedule.Background  = Brush 'Transparent'
        $TabSchedule.Foreground  = Brush '#8A8A90'
        Render-Month
    } else {
        $MonthView.Visibility    = 'Collapsed'
        $ScheduleView.Visibility = 'Visible'
        $TabSchedule.Background  = Brush '#33FFFFFF'
        $TabSchedule.Foreground  = Brush '#F5F5F7'
        $TabMonth.Background     = Brush 'Transparent'
        $TabMonth.Foreground     = Brush '#8A8A90'
    }

    # the notes block only belongs to the schedule view
    $notesVis = $(if ($name -eq 'schedule') { 'Visible' } else { 'Collapsed' })
    $NoteHeaderRow.Visibility = $notesVis
    $NoteAreaRow.Visibility   = $notesVis

    # collapse the fixed 190px row too, otherwise it squeezes the month view
    if ($name -eq 'schedule') {
        $PanelGrid.RowDefinitions[3].Height = New-Object System.Windows.GridLength(190)
    } else {
        $PanelGrid.RowDefinitions[3].Height = New-Object System.Windows.GridLength(0)
    }
}

$script:MvSelected = (Get-Date).ToString('yyyy-MM-dd')

function Render-MvDayEvents {
    $MvDayEvents.Children.Clear()
    $key = $script:MvSelected

    $dt = [datetime]::ParseExact($key, 'yyyy-MM-dd', $null)
    $MvDayTitle.Text = "$($MON[$dt.Month-1])$($dt.Day)日 $($WD[[int]$dt.DayOfWeek])"

    $evs = @($script:Events | Where-Object { $_.date -eq $key } | Sort-Object start)
    foreach ($e in $evs) {
        $row = New-Object System.Windows.Controls.Border
        $row.CornerRadius = '7'
        $row.Background = Brush '#14FFFFFF'
        $row.Margin = '0,0,0,4'
        $row.Padding = '9,6,9,7'

        $g = New-Object System.Windows.Controls.Grid
        $c0 = New-Object System.Windows.Controls.ColumnDefinition
        $c0.Width = '4'
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $g.ColumnDefinitions.Add($c0)
        $g.ColumnDefinitions.Add($c1)

        $bar = New-Object System.Windows.Controls.Border
        $bar.Background = Brush (ColorFor $e.title)
        $bar.CornerRadius = '2,0,0,2'
        [System.Windows.Controls.Grid]::SetColumn($bar, 0)
        $g.Children.Add($bar) | Out-Null

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Margin = '9,0,0,0'
        [System.Windows.Controls.Grid]::SetColumn($sp, 1)

        $tt = New-Object System.Windows.Controls.TextBlock
        $tt.Text = "$($e.time) " + (ShortName $e.title)
        $tt.Foreground = Brush '#F5F5F7'
        $tt.FontSize = 11.5
        $tt.TextTrimming = 'CharacterEllipsis'
        $sp.Children.Add($tt) | Out-Null

        if ($e.location) {
            $tl = New-Object System.Windows.Controls.TextBlock
            $tl.Text = $e.location
            $tl.Foreground = Brush '#5E5E64'
            $tl.FontSize = 9.5
            $tl.TextTrimming = 'CharacterEllipsis'
            $tl.Margin = '0,2,0,0'
            $sp.Children.Add($tl) | Out-Null
        }
        $g.Children.Add($sp) | Out-Null
        $row.Child = $g
        $MvDayEvents.Children.Add($row) | Out-Null
    }

    $dueNotes = @($script:Notes | Where-Object { -not $_.done -and $_.due -and ([string]$_.due) -eq $key })
    foreach ($n in $dueNotes) {
        $row = New-Object System.Windows.Controls.Border
        $row.CornerRadius = '7'
        $row.Background = Brush '#1FFF9F0A'
        $row.Margin = '0,0,0,4'
        $row.Padding = '9,6,9,7'
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = 'Horizontal'
        $tag = New-Object System.Windows.Controls.TextBlock
        $tag.Text = '截止'
        $tag.Foreground = Brush '#FF9F0A'
        $tag.FontSize = 9.5
        $tag.VerticalAlignment = 'Center'
        $tag.Margin = '0,0,7,0'
        $sp.Children.Add($tag) | Out-Null
        $tx = New-Object System.Windows.Controls.TextBlock
        $tx.Text = $n.text
        $tx.Foreground = Brush '#E8E8EC'
        $tx.FontSize = 11.5
        $sp.Children.Add($tx) | Out-Null
        $row.Child = $sp
        $MvDayEvents.Children.Add($row) | Out-Null
    }

    if ($evs.Count -eq 0 -and $dueNotes.Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '这天没有安排'
        $t.Foreground = Brush '#5E5E64'
        $t.FontSize = 11
        $t.Margin = '2,4,0,0'
        $MvDayEvents.Children.Add($t) | Out-Null
    }
}

function Render-Month {
    $MvTitle.Text = "$($script:MvYear) 年 $($script:MvMonth) 月"
    $MvGrid.Children.Clear()

    $first  = New-Object DateTime($script:MvYear, $script:MvMonth, 1)
    $offset = [int]$first.DayOfWeek
    $days   = [DateTime]::DaysInMonth($script:MvYear, $script:MvMonth)
    $todayKey = (Get-Date).ToString('yyyy-MM-dd')

    $dueSet = @{}
    foreach ($n in $script:Notes) {
        if (-not $n.done -and $n.due) { $dueSet[[string]$n.due] = $true }
    }

    for ($i = 0; $i -lt $offset; $i++) {
        $MvGrid.Children.Add((New-Object System.Windows.Controls.Border)) | Out-Null
    }

    for ($d = 1; $d -le $days; $d++) {
        $key = "$($script:MvYear)-$($script:MvMonth.ToString('00'))-$($d.ToString('00'))"

        $cell = New-Object System.Windows.Controls.Border
        $cell.CornerRadius = '7'
        $cell.Margin = '1'
        $cell.Background = Brush 'Transparent'
        $cell.Cursor = 'Hand'
        $cell.Tag = $key

        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.VerticalAlignment = 'Center'
        $sp.HorizontalAlignment = 'Center'

        $ring = New-Object System.Windows.Controls.Border
        $ring.Width = 24
        $ring.Height = 24
        $ring.CornerRadius = '12'
        $ring.BorderThickness = '1.5'
        $ring.BorderBrush = Brush $(if ($key -eq $script:MvSelected) { '#FF453A' } else { 'Transparent' })
        $ring.HorizontalAlignment = 'Center'

        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = [string]$d
        $t.FontSize = 10.5
        $t.HorizontalAlignment = 'Center'
        $t.VerticalAlignment = 'Center'
        if ($key -eq $todayKey) {
            $t.Foreground = Brush '#FF453A'
            $t.FontWeight = 'SemiBold'
        } else {
            $t.Foreground = Brush '#D8D8DC'
        }
        $ring.Child = $t
        $sp.Children.Add($ring) | Out-Null

        $dots = New-Object System.Windows.Controls.StackPanel
        $dots.Orientation = 'Horizontal'
        $dots.HorizontalAlignment = 'Center'
        $dots.Margin = '0,3,0,0'
        $dots.Height = 5

        $evs = @($script:Events | Where-Object { $_.date -eq $key })
        if ($evs.Count -gt 0) {
            $seen = @{}
            foreach ($e in $evs) {
                $c = ColorFor $e.title
                if ($seen.ContainsKey($c)) { continue }
                $seen[$c] = 1
                $dot = New-Object System.Windows.Shapes.Ellipse
                $dot.Width = 4; $dot.Height = 4
                $dot.Fill = Brush $c
                $dot.Margin = '0,0,2,0'
                $dots.Children.Add($dot) | Out-Null
                if ($dots.Children.Count -ge 3) { break }
            }
        }
        if ($dueSet.ContainsKey($key)) {
            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = 4; $dot.Height = 4
            $dot.Fill = Brush '#FF9F0A'
            $dot.Margin = '0,0,2,0'
            $dots.Children.Add($dot) | Out-Null
        }
        if ($dots.Children.Count -gt 0) { $sp.Children.Add($dots) | Out-Null }

        $cell.Child = $sp
        $cell.Add_MouseLeftButtonUp({
            param($s, $ev)
            $script:MvSelected = [string]$s.Tag
            Render-Month
        })
        $MvGrid.Children.Add($cell) | Out-Null
    }

    Render-MvDayEvents
}

$TabSchedule.Add_Click({ Switch-Tab 'schedule' })
$TabMonth.Add_Click({ Switch-Tab 'month' })

$MvPrev.Add_Click({
    $script:MvMonth--
    if ($script:MvMonth -lt 1) { $script:MvMonth = 12; $script:MvYear-- }
    Render-Month
})
$MvNext.Add_Click({
    $script:MvMonth++
    if ($script:MvMonth -gt 12) { $script:MvMonth = 1; $script:MvYear++ }
    Render-Month
})

# ---------------------------------------------------------------- settings window wiring
$SetCloseBtn.Add_Click({ $sw.Hide() })

$SetHeader.Add_MouseLeftButtonDown({ $sw.DragMove() })

$SetUrlInput.Add_TextChanged({
    if ($SetUrlInput.Text) { $SetUrlHint.Visibility = 'Collapsed' } else { $SetUrlHint.Visibility = 'Visible' }
})

$SetImportBtn.Add_Click({ Import-IcsSource })

$SettingsBtn.Add_Click({
    if (-not $sw.IsVisible) {
        $setLeft = $w.Left - 470 - 12
        if ($setLeft -lt $wa.Left + 8) { $setLeft = $wa.Left + 8 }
        $sw.Left = $setLeft
        $sw.Top  = $w.Top
        $sw.Show()
    } else {
        $sw.Activate()
    }
    Render-Settings
})

# ---------------------------------------------------------------- lock / layer
$script:Locked = $false
$script:OnTop  = $true

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class OrbWin32 {
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
}
"@ -ErrorAction SilentlyContinue

function Update-LockUI {
    if ($script:Locked) {
        $LockIcon.Text = [string][char]0xE72E
        $LockIcon.Foreground = Brush '#FF9F0A'
    } else {
        $LockIcon.Text = [string][char]0xE785
        $LockIcon.Foreground = Brush '#7A7A80'
    }
}

function Update-LayerUI {
    $LayerIcon.Foreground = Brush $(if ($script:OnTop) { '#0A84FF' } else { '#5E5E64' })
}

$LockBtn.Add_MouseLeftButtonDown({
    param($s, $ev)
    $ev.Handled = $true
})
$LayerBtn.Add_MouseLeftButtonDown({
    param($s, $ev)
    $ev.Handled = $true
})

$LockBtn.Add_MouseLeftButtonUp({
    param($s, $ev)
    $ev.Handled = $true
    $script:Locked = -not $script:Locked
    Update-LockUI
})

$LayerBtn.Add_MouseLeftButtonUp({
    param($s, $ev)
    $ev.Handled = $true
    $script:OnTop = -not $script:OnTop
    if ($script:OnTop) {
        $w.Topmost = $true
    } else {
        $w.Topmost = $false
        try {
            $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
            [void][OrbWin32]::SetWindowPos($hwnd, [IntPtr]1, 0, 0, 0, 0, 0x0003)
        } catch {}
    }
    Update-LayerUI
})

Update-LockUI
Update-LayerUI

# ---------------------------------------------------------------- render classes
$WD = @('周日','周一','周二','周三','周四','周五','周六')
$MON = @('1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月')

function Add-DayHeader($dayKey) {
    $dt = [datetime]::ParseExact($dayKey, 'yyyy-MM-dd', $null)
    $todayKey = (Get-Date).ToString('yyyy-MM-dd')
    $diff = ($dt.Date - (Get-Date).Date).Days
    $label = switch ($diff) { 0 { '今天' } 1 { '明天' } 2 { '后天' } default { '' } }

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'
    $sp.Margin = '4,10,0,7'

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.Text = "$($MON[$dt.Month-1])$($dt.Day)日 $($WD[[int]$dt.DayOfWeek])"
    $t1.Foreground = Brush '#F5F5F7'
    $t1.FontSize = 12.5
    $t1.FontWeight = 'SemiBold'
    $sp.Children.Add($t1) | Out-Null

    if ($label) {
        $t2 = New-Object System.Windows.Controls.TextBlock
        $t2.Text = $label
        $t2.Foreground = Brush '#8A8A90'
        $t2.FontSize = 11
        $t2.Margin = '8,1,0,0'
        $sp.Children.Add($t2) | Out-Null
    }
    $ClassList.Children.Add($sp) | Out-Null
}

function Add-ClassRow($e) {
    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = '9'
    $border.Background = Brush '#14FFFFFF'
    $border.Margin = '0,0,0,5'

    $grid = New-Object System.Windows.Controls.Grid
    $col0 = New-Object System.Windows.Controls.ColumnDefinition
    $col0.Width = '4'
    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $grid.ColumnDefinitions.Add($col0)
    $grid.ColumnDefinitions.Add($col1)

    $bar = New-Object System.Windows.Controls.Border
    $bar.Background = Brush (ColorFor $e.title)
    $bar.CornerRadius = '2,0,0,2'
    [System.Windows.Controls.Grid]::SetColumn($bar, 0)
    $grid.Children.Add($bar) | Out-Null

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = '11,8,10,9'
    [System.Windows.Controls.Grid]::SetColumn($sp, 1)

    $tc = New-Object System.Windows.Controls.TextBlock
    $endT = ''
    if ($e.end) { $endT = ' – ' + ([datetime]$e.end).ToString('HH:mm') }
    $tc.Text = "$($e.time)$endT"
    $tc.Foreground = Brush '#8A8A90'
    $tc.FontSize = 10.5
    $sp.Children.Add($tc) | Out-Null

    $tt = New-Object System.Windows.Controls.TextBlock
    $tt.Text = ShortName $e.title
    $tt.Foreground = Brush '#F5F5F7'
    $tt.FontSize = 12.5
    $tt.FontWeight = 'SemiBold'
    $tt.TextTrimming = 'CharacterEllipsis'
    $tt.Margin = '0,2,0,0'
    $sp.Children.Add($tt) | Out-Null

    if ($e.location) {
        $tl = New-Object System.Windows.Controls.TextBlock
        $tl.Text = $e.location
        $tl.Foreground = Brush '#5E5E64'
        $tl.FontSize = 10.5
        $tl.TextTrimming = 'CharacterEllipsis'
        $tl.Margin = '0,2,0,0'
        $sp.Children.Add($tl) | Out-Null
    }

    $grid.Children.Add($sp) | Out-Null
    $border.Child = $grid
    $ClassList.Children.Add($border) | Out-Null
}

function Add-DeadlineRow($n) {
    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = '9'
    $border.Background = Brush '#1FFF9F0A'
    $border.Margin = '0,0,0,5'
    $border.Padding = '11,7,10,8'

    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'

    $tag = New-Object System.Windows.Controls.TextBlock
    $tag.Text = '截止'
    $tag.Foreground = Brush '#FF9F0A'
    $tag.FontSize = 9.5
    $tag.VerticalAlignment = 'Center'
    $tag.Margin = '0,0,7,0'
    $sp.Children.Add($tag) | Out-Null

    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $n.text
    $t.Foreground = Brush '#E8E8EC'
    $t.FontSize = 12
    $t.TextTrimming = 'CharacterEllipsis'
    $t.VerticalAlignment = 'Center'
    $sp.Children.Add($t) | Out-Null

    $border.Child = $sp
    $border.Tag = $n.id
    $border.Cursor = 'Hand'
    $border.Add_MouseLeftButtonUp({
        param($s, $ev)
        $id = $s.Tag
        foreach ($item in $script:Notes) {
            if ($item.id -eq $id) { $item.done = -not $item.done; break }
        }
        Save-Notes $script:Notes
        Render-Classes
        Render-Notes
    })
    $ClassList.Children.Add($border) | Out-Null
}

function Render-Classes {
    $ClassList.Children.Clear()
    $todayKey = (Get-Date).ToString('yyyy-MM-dd')
    $up = @($script:Events | Where-Object { $_.date -ge $todayKey } | Sort-Object start)

    if ($up.Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '这一学期没有更多课程了'
        $t.Foreground = Brush '#5E5E64'
        $t.FontSize = 12
        $t.Margin = '6,24,0,0'
        $ClassList.Children.Add($t) | Out-Null
        return
    }

    $byDay = [ordered]@{}
    foreach ($e in $up) {
        if (-not $byDay.Contains($e.date)) { $byDay[$e.date] = @() }
        $byDay[$e.date] = @($byDay[$e.date]) + $e
    }
    $keys = @($byDay.Keys | Sort-Object | Select-Object -First 5)
    foreach ($k in $keys) {
        Add-DayHeader $k
        foreach ($e in $byDay[$k]) { Add-ClassRow $e }
        $dueNotes = @($script:Notes | Where-Object { -not $_.done -and $_.due -and ([string]$_.due) -eq $k })
        foreach ($n in $dueNotes) { Add-DeadlineRow $n }
    }

    # deadlines today even if there is no class today
    $todayKey2 = (Get-Date).ToString('yyyy-MM-dd')
    if (-not $byDay.Contains($todayKey2)) {
        $todayDue = @($script:Notes | Where-Object { -not $_.done -and $_.due -and ([string]$_.due) -eq $todayKey2 })
        if ($todayDue.Count -gt 0) {
            Add-DayHeader $todayKey2
            foreach ($n in $todayDue) { Add-DeadlineRow $n }
        }
    }
}

# ---------------------------------------------------------------- date picker
function Build-Calendar {
    $CalGrid.Children.Clear()
    $CalTitle.Text = "$($script:CalYear) 年 $($script:CalMonth) 月"

    foreach ($wd in @('日','一','二','三','四','五','六')) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $wd
        $t.Foreground = Brush '#5E5E64'
        $t.FontSize = 9.5
        $t.HorizontalAlignment = 'Center'
        $t.VerticalAlignment = 'Center'
        $t.Margin = '0,0,0,5'
        $CalGrid.Children.Add($t) | Out-Null
    }

    $first = New-Object DateTime($script:CalYear, $script:CalMonth, 1)
    $offset = [int]$first.DayOfWeek
    $days = [DateTime]::DaysInMonth($script:CalYear, $script:CalMonth)
    $todayKey = (Get-Date).ToString('yyyy-MM-dd')

    for ($i = 0; $i -lt $offset; $i++) {
        $CalGrid.Children.Add((New-Object System.Windows.Controls.Border)) | Out-Null
    }

    for ($d = 1; $d -le $days; $d++) {
        $key = "$($script:CalYear)-$($script:CalMonth.ToString('00'))-$($d.ToString('00'))"
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = [string]$d
        $btn.Width = 26
        $btn.Height = 24
        $btn.Margin = '1'
        $btn.Padding = '0'
        $btn.BorderThickness = '0'
        $btn.Background = Brush 'Transparent'
        $btn.Foreground = Brush '#D8D8DC'
        $btn.FontSize = 10.5
        $btn.Cursor = 'Hand'
        $btn.Tag = $key

        if ($key -eq $todayKey) {
            $btn.Foreground = Brush '#FF453A'
            $btn.FontWeight = 'SemiBold'
        }
        if ($script:DueDate -and $key -eq $script:DueDate) {
            $btn.Background = Brush '#0A84FF'
            $btn.Foreground = Brush '#FFFFFF'
        }

        $btn.Add_Click({
            param($s, $ev)
            $script:DueDate = [string]$s.Tag
            $DueText.Text = '截止 ' + (Format-DueShort $script:DueDate)
            $DueChip.Visibility = 'Visible'
            $CalPopup.IsOpen = $false
        })
        $CalGrid.Children.Add($btn) | Out-Null
    }
}

# ---------------------------------------------------------------- render notes
function Render-Notes {
    $NoteList.Children.Clear()

    # notes WITH a due date live in the class list, not here
    $pending = @($Notes | Where-Object { -not $_.due })
    $active  = @($pending | Where-Object { -not $_.done })
    $NoteCount.Text = if ($active.Count -gt 0) { "$($active.Count) 条待办" } else { '' }

    if ($pending.Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '还没有备忘'
        $t.Foreground = Brush '#5E5E64'
        $t.FontSize = 11.5
        $t.Margin = '5,6,0,0'
        $NoteList.Children.Add($t) | Out-Null
        return
    }

    $sorted = @($pending | Sort-Object `
        @{Expression={[bool]$_.done}},
        @{Expression={ if ($_.due) { [string]$_.due } else { '9999-99-99' } }},
        @{Expression={$_.created}; Descending=$true})

    foreach ($n in $sorted) {
        $row = New-Object System.Windows.Controls.Border
        $row.CornerRadius = '7'
        $row.Background = Brush '#10FFFFFF'
        $row.Margin = '0,0,0,4'
        $row.Padding = '9,7,9,7'
        $row.Tag = $n.id

        $g = New-Object System.Windows.Controls.Grid
        $c0 = New-Object System.Windows.Controls.ColumnDefinition
        $c0.Width = '22'
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $c2 = New-Object System.Windows.Controls.ColumnDefinition
        $c2.Width = '20'
        $g.ColumnDefinitions.Add($c0)
        $g.ColumnDefinitions.Add($c1)
        $g.ColumnDefinitions.Add($c2)

        $box = New-Object System.Windows.Controls.Border
        $box.Width = 14; $box.Height = 14
        $box.CornerRadius = '7'
        $box.BorderThickness = '1.3'
        $box.HorizontalAlignment = 'Left'
        $box.VerticalAlignment = 'Center'
        if ($n.done) {
            $box.Background = Brush '#30D158'
            $box.BorderBrush = Brush '#30D158'
        } else {
            $box.Background = Brush 'Transparent'
            $box.BorderBrush = Brush '#6E6E73'
        }
        [System.Windows.Controls.Grid]::SetColumn($box, 0)
        $g.Children.Add($box) | Out-Null

        $textCol = New-Object System.Windows.Controls.StackPanel
        $textCol.VerticalAlignment = 'Center'

        $txt = New-Object System.Windows.Controls.TextBlock
        $txt.Text = $n.text
        $txt.FontSize = 12
        $txt.TextWrapping = 'Wrap'
        $txt.Foreground = Brush $(if ($n.done) { '#5E5E64' } else { '#E8E8EC' })
        if ($n.done) { $txt.TextDecorations = [System.Windows.TextDecorations]::Strikethrough }
        $textCol.Children.Add($txt) | Out-Null

        if ($n.due) {
            $dueStr = [string]$n.due
            $todayKey = (Get-Date).ToString('yyyy-MM-dd')
            $dt = New-Object System.Windows.Controls.TextBlock
            $dt.FontSize = 9.5
            $dt.Margin = '0,2,0,0'
            if ($dueStr -lt $todayKey) {
                $dt.Text = '逾期 · ' + (Format-DueShort $dueStr)
                $dt.Foreground = Brush '#FF453A'
            } elseif ($dueStr -eq $todayKey) {
                $dt.Text = '今天到期'
                $dt.Foreground = Brush '#FF9F0A'
            } else {
                $dt.Text = (Format-DueShort $dueStr) + ' 到期'
                $dt.Foreground = Brush '#6E6E73'
            }
            $textCol.Children.Add($dt) | Out-Null
        }

        [System.Windows.Controls.Grid]::SetColumn($textCol, 1)
        $g.Children.Add($textCol) | Out-Null

        $x = New-Object System.Windows.Controls.TextBlock
        $x.Text = '✕'
        $x.FontSize = 11
        $x.Foreground = Brush '#6E6E73'
        $x.Cursor = 'Hand'
        $x.HorizontalAlignment = 'Right'
        $x.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($x, 2)
        $g.Children.Add($x) | Out-Null

        $row.Child = $g
        $row.Cursor = 'Hand'

        $row.Add_MouseLeftButtonUp({
            param($s, $ev)
            $id = $s.Tag
            foreach ($item in $script:Notes) {
                if ($item.id -eq $id) { $item.done = -not $item.done; break }
            }
            Save-Notes $script:Notes
            Render-Notes
        })

        $x.Add_MouseLeftButtonUp({
            param($s, $ev)
            $ev.Handled = $true
            $id = $s.Tag
            $script:Notes = @($script:Notes | Where-Object { $_.id -ne $id })
            Save-Notes $script:Notes
            Render-Notes
        })
        $x.Tag = $n.id

        $NoteList.Children.Add($row) | Out-Null
    }
}

# ---------------------------------------------------------------- settings / import
function Render-Settings {
    $SetSourceList.Children.Clear()
    if (@($script:Sources).Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '还没有导入任何日历'
        $t.Foreground = Brush '#5E5E64'
        $t.FontSize = 11
        $t.Margin = '2,6,0,0'
        $SetSourceList.Children.Add($t) | Out-Null
        return
    }

    foreach ($s in $script:Sources) {
        $row = New-Object System.Windows.Controls.Border
        $row.CornerRadius = '8'
        $row.Background = Brush '#12FFFFFF'
        $row.Margin = '0,0,0,5'
        $row.Padding = '10,8,10,8'

        $g = New-Object System.Windows.Controls.Grid
        $c0 = New-Object System.Windows.Controls.ColumnDefinition
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $c1.Width = '22'
        $g.ColumnDefinitions.Add($c0)
        $g.ColumnDefinitions.Add($c1)

        $sp = New-Object System.Windows.Controls.StackPanel

        $n1 = New-Object System.Windows.Controls.TextBlock
        $n1.Text = [string]$s.name
        $n1.Foreground = Brush '#F5F5F7'
        $n1.FontSize = 11.5
        $n1.TextTrimming = 'CharacterEllipsis'
        $sp.Children.Add($n1) | Out-Null

        $cnt = 0
        if ($s.events) { $cnt = @($s.events).Count }
        $n2 = New-Object System.Windows.Controls.TextBlock
        $n2.Text = "$cnt 个事件"
        $n2.Foreground = Brush '#6E6E73'
        $n2.FontSize = 9.5
        $n2.Margin = '0,2,0,0'
        $sp.Children.Add($n2) | Out-Null

        [System.Windows.Controls.Grid]::SetColumn($sp, 0)
        $g.Children.Add($sp) | Out-Null

        $x = New-Object System.Windows.Controls.TextBlock
        $x.Text = '✕'
        $x.FontSize = 11
        $x.Foreground = Brush '#6E6E73'
        $x.Cursor = 'Hand'
        $x.HorizontalAlignment = 'Right'
        $x.VerticalAlignment = 'Center'
        $x.Tag = [string]$s.id
        [System.Windows.Controls.Grid]::SetColumn($x, 1)
        $g.Children.Add($x) | Out-Null

        $x.Add_MouseLeftButtonUp({
            param($snd, $evt)
            $id = [string]$snd.Tag
            $script:Sources = @($script:Sources | Where-Object { [string]$_.id -ne $id })
            Save-Sources $script:Sources
            Rebuild-Events
            Render-Settings
            Render-Classes
            Render-Month
        })

        $row.Child = $g
        $SetSourceList.Children.Add($row) | Out-Null
    }
}

function Import-IcsSource {
    $url = [string]$SetUrlInput.Text
    if (-not $url -or -not $url.Trim()) {
        $SetStatus.Text = '先粘贴一个链接'
        $SetStatus.Foreground = Brush '#FF9F0A'
        return
    }
    $url = $url.Trim()
    if ($url -like 'webcal://*') { $url = 'https://' + $url.Substring(9) }
    if ($url -notmatch '^https?://') {
        $SetStatus.Text = '只支持 http / https / webcal 链接'
        $SetStatus.Foreground = Brush '#FF453A'
        return
    }

    $SetStatus.Text = '正在下载…'
    $SetStatus.Foreground = Brush '#8A8A90'

    try {
        $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 25
        $text = [string]$resp.Content
        if ($text -notmatch 'BEGIN:VCALENDAR') {
            $SetStatus.Text = '这个链接不是 .ics 日历'
            $SetStatus.Foreground = Brush '#FF453A'
            return
        }

        $evs = @(Parse-IcsText $text)
        if ($evs.Count -eq 0) {
            $SetStatus.Text = '下载成功，但里面没有可用事件'
            $SetStatus.Foreground = Brush '#FF9F0A'
            return
        }

        $name = ''
        try {
            $u = [uri]$url
            $name = [System.IO.Path]::GetFileNameWithoutExtension($u.AbsolutePath)
        } catch {}
        if (-not $name) { $name = '导入的日历' }

        $src = [pscustomobject]@{
            id     = [guid]::NewGuid().ToString()
            name   = $name
            url    = $url
            added  = (Get-Date).ToString('o')
            events = $evs
        }
        $script:Sources = @($script:Sources) + $src
        Save-Sources $script:Sources
        Rebuild-Events

        $SetStatus.Text = "已导入 $($evs.Count) 个事件（$name）"
        $SetStatus.Foreground = Brush '#30D158'
        $SetUrlInput.Text = ''
        $SetUrlHint.Visibility = 'Visible'
        Render-Settings
        Render-Classes
        Render-Month
    } catch {
        $SetStatus.Text = '下载失败：' + $_.Exception.Message
        $SetStatus.Foreground = Brush '#FF453A'
    }
}

# ---------------------------------------------------------------- note input
function Add-Note {
    $text = $NoteInput.Text
    if (-not $text -or -not $text.Trim()) { return }
    try {
        $item = [pscustomobject]@{
            id      = [guid]::NewGuid().ToString()
            text    = $text.Trim()
            done    = $false
            created = (Get-Date).ToString('o')
            due     = $(if ($script:DueDate) { [string]$script:DueDate } else { $null })
        }
        $script:Notes = @($script:Notes) + $item
        Save-Notes $script:Notes
        $NoteInput.Text = ''
        $NoteHint.Visibility = 'Visible'
        $script:DueDate = $null
        $DueChip.Visibility = 'Collapsed'
        Render-Notes
    } catch {
        Write-Err 'add-note'
    }
}

$NoteInput.Add_TextChanged({
    if ($NoteInput.Text) { $NoteHint.Visibility = 'Collapsed' } else { $NoteHint.Visibility = 'Visible' }
})

$AddBtn.Add_Click({ Add-Note })

$CalPopup.PlacementTarget = $DueBtn
$CalPopup.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Top

$DueBtn.Add_Click({
    $script:CalYear  = (Get-Date).Year
    $script:CalMonth = (Get-Date).Month
    Build-Calendar
    $CalPopup.IsOpen = $true
})

$CalPrev.Add_Click({
    $script:CalMonth--
    if ($script:CalMonth -lt 1) { $script:CalMonth = 12; $script:CalYear-- }
    Build-Calendar
})

$CalNext.Add_Click({
    $script:CalMonth++
    if ($script:CalMonth -gt 12) { $script:CalMonth = 1; $script:CalYear++ }
    Build-Calendar
})

$DueClear.Add_MouseLeftButtonUp({
    $script:DueDate = $null
    $DueChip.Visibility = 'Collapsed'
})

$NoteInput.Add_PreviewKeyDown({
    param($s, $ev)
    if ($ev.Key -eq [System.Windows.Input.Key]::Enter) {
        Add-Note
        $ev.Handled = $true
    }
})

# ---------------------------------------------------------------- expand / collapse
$script:IsOpen = $false

function Write-Err([string]$where) {
    try {
        $msg = (Get-Date).ToString('s') + " [" + $where + "] " + $_.Exception.ToString() + "`r`n"
        [System.IO.File]::AppendAllText((Join-Path $Base 'error.log'), $msg, (New-Object System.Text.UTF8Encoding($false)))
    } catch {}
}

function Expand-Orb {
    if ($script:IsOpen) { return }

    try {
        Render-Classes
        Render-Notes
    } catch {
        Write-Err 'render'
        return
    }

    $script:IsOpen = $true
    $right = $w.Left + 470
    $w.Width = 380
    $w.Height = 560
    $newLeft = $right - 380

    # clamp only when the panel would fall outside every physical screen
    $fits = $false
    foreach ($sc in [System.Windows.Forms.Screen]::AllScreens) {
        $b = $sc.Bounds
        if ($newLeft -ge ($b.Left - 10) -and $newLeft -le ($b.Right - 120)) { $fits = $true; break }
    }
    if (-not $fits) { $newLeft = $wa.Left + 8 }
    $w.Left = $newLeft

    $OrbView.Visibility = 'Collapsed'
    $PanelView.Visibility = 'Visible'
}

function Collapse-Orb {
    if (-not $script:IsOpen) { return }
    $script:IsOpen = $false

    $right = $w.Left + $w.Width
    $w.Width = 470
    $w.Height = 76
    $w.Left = $right - 470

    $PanelView.Visibility = 'Collapsed'
    $OrbView.Visibility = 'Visible'
    Save-State
    Update-Orb
}

$CloseBtn.Add_Click({ Collapse-Orb })

$OrbView.Add_MouseLeftButtonDown({
    if ($script:Locked) { return }
    $script:preLeft = $w.Left
    $script:preTop  = $w.Top
    $w.DragMove()
    if ([math]::Abs($w.Left - $script:preLeft) -lt 3 -and [math]::Abs($w.Top - $script:preTop) -lt 3) {
        Expand-Orb
    } else {
        Save-State
    }
})

# ---------------------------------------------------------------- tray
$ni = New-Object System.Windows.Forms.NotifyIcon
$iconPath = Join-Path $Base 'assets\beacon.ico'
if (-not (Test-Path $iconPath)) { $iconPath = Join-Path $Base 'app\assets\beacon.ico' }
if (Test-Path $iconPath) {
    try { $ni.Icon = New-Object System.Drawing.Icon($iconPath) }
    catch { $ni.Icon = [System.Drawing.SystemIcons]::Application }
} else {
    $ni.Icon = [System.Drawing.SystemIcons]::Application
}
$ni.Text = 'beacon'
$ni.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miShow = New-Object System.Windows.Forms.ToolStripMenuItem '展开'
$miShow.Add_Click({
    if (-not $w.IsVisible) { $w.Show() }
    Expand-Orb
})
$miQuit = New-Object System.Windows.Forms.ToolStripMenuItem '退出'
$miQuit.Add_Click({
    Save-State
    $ni.Visible = $false
    $w.Close()
})
$menu.Items.Add($miShow) | Out-Null
$menu.Items.Add($miQuit) | Out-Null
$ni.ContextMenuStrip = $menu

$ni.Add_MouseClick({
    param($s, $ev)
    if ($ev.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        if ($script:IsOpen) { Collapse-Orb } else { Expand-Orb }
    }
})

# ---------------------------------------------------------------- run
Update-Orb
$w.Topmost = $true

if ($TestExpand) {
    $w.Add_ContentRendered({ Expand-Orb })
}

$app = New-Object System.Windows.Application
$app.Add_Exit({
    try { $ni.Visible = $false; $ni.Dispose() } catch {}
    try { $script:OrbMutex.ReleaseMutex() } catch {}
})

$w.Add_ContentRendered({
    try {
        $p0 = New-Object System.Windows.Point(0,0)
        $po = $OrbView.TranslatePoint($p0, $w)
        $pl = $LockBtn.TranslatePoint($p0, $w)
        $msg = "window=$($w.ActualWidth)x$($w.ActualHeight)`r`n" +
               "orb=$($OrbView.ActualWidth)x$($OrbView.ActualHeight) at $($po.X),$($po.Y)`r`n" +
               "lock=$($LockBtn.ActualWidth)x$($LockBtn.ActualHeight) at $($pl.X),$($pl.Y)"
        [System.IO.File]::WriteAllText((Join-Path $Base 'layout.txt'), $msg, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        [System.IO.File]::WriteAllText((Join-Path $Base 'layout.txt'), "ERR: $($_.Exception.Message)", (New-Object System.Text.UTF8Encoding($false)))
    }
})

# metrics refresh every 2s
$metricTimer = New-Object System.Windows.Threading.DispatcherTimer
$metricTimer.Interval = [TimeSpan]::FromSeconds(2)
$metricTimer.Add_Tick({ Update-Orb })
$metricTimer.Start()

# schedule refresh every 5min
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMinutes(5)
$timer.Add_Tick({
    if ($script:IsOpen) { Render-Classes }
})
$timer.Start()

# ---------------------------------------------------------------- 22:00 auto refresh of imported sources
function Refresh-AllSources {
    if (@($script:Sources).Count -eq 0) { return }
    $newList = @()
    foreach ($s in $script:Sources) {
        try {
            $resp = Invoke-WebRequest -Uri $s.url -UseBasicParsing -TimeoutSec 25
            $text = [string]$resp.Content
            if ($text -match 'BEGIN:VCALENDAR') {
                $evs = @(Parse-IcsText $text)
                if ($evs.Count -gt 0) { $s.events = $evs }
            }
        } catch {}
        $newList += $s
    }
    $script:Sources = $newList
    Save-Sources $script:Sources
    Rebuild-Events
    if ($script:IsOpen) { Render-Classes; Render-Month }
}

$script:LastAutoRefresh = ''
$refreshTimer = New-Object System.Windows.Threading.DispatcherTimer
$refreshTimer.Interval = [TimeSpan]::FromSeconds(60)
$refreshTimer.Add_Tick({
    $n = Get-Date
    if ($n.Hour -eq 22 -and $n.Minute -eq 0) {
        $d = $n.ToString('yyyy-MM-dd')
        if ($script:LastAutoRefresh -ne $d) {
            $script:LastAutoRefresh = $d
            Refresh-AllSources
        }
    }
})
$refreshTimer.Start()

$app.Run($w) | Out-Null
