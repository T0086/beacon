param(
    [switch]$AutoStart,
    [switch]$TestExpand,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# ---------------------------------------------------------------- single instance guard
# Without this, every restart leaves a ghost tray icon behind and a second copy
# can be launched by both the scheduled task and a manual double-click.
if (-not $SelfTest) {
    $script:OrbMutex = New-Object System.Threading.Mutex($false, 'Global\Beacon_SingleInstance')
    if (-not $script:OrbMutex.WaitOne(0)) {
        exit
    }
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
if (-not $Base) { $Base = (Get-Location).Path }
# optional external timetable file; if it is missing the schedule pane just shows empty
$SchedulePath = Join-Path (Split-Path $Base -Parent) 'desktop-calendar\schedule.json'
$NotesPath    = Join-Path $Base 'notes.json'
$StatePath    = Join-Path $Base 'state.json'
$CtlDir       = Join-Path $Base 'ctl'
if (-not (Test-Path $CtlDir)) { New-Item -ItemType Directory -Path $CtlDir -Force | Out-Null }

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

function Reload-Notes {
    # 外部（beacon-sync / 别处写入）改过 notes.json 时内存里那份是旧的；
    # 开面板、刷新、定时重绘之前从这里重读一次。读不懂就保持原样，不把内存清空。
    try {
        if (-not (Test-Path $NotesPath)) { return }
        $raw = [System.IO.File]::ReadAllText($NotesPath, [System.Text.Encoding]::UTF8)
        if (-not $raw -or -not $raw.Trim()) { return }
        $o = $raw | ConvertFrom-Json
        if ($null -eq $o) { return }
        $items = @($o.items | Where-Object { $null -ne $_ })
        foreach ($it in $items) {
            if ($null -ne $it.due -and $it.due -isnot [string]) {
                $it.due = ([datetime]$it.due).ToString('yyyy-MM-dd')
            }
        }
        $script:Notes = $items
    } catch {}
}

function Get-AgendaDays {
    # 日程列表取哪几天：有课的日期 ∪ 有截止日的备忘日期（都 >= 今天），排序取前 5 天。
    # 原来只按「有课的日子」取 key，于是「今天只有备忘」「备忘落在没课那天」整块被吃掉。
    $todayKey = (Get-Date).ToString('yyyy-MM-dd')
    $days = [ordered]@{}
    $cls  = @{}
    $nts  = @{}

    foreach ($e in @($script:Events)) {
        if ($null -eq $e -or -not $e.date) { continue }
        $d = [string]$e.date
        if ($d -lt $todayKey) { continue }
        if (-not $days.Contains($d)) { $days[$d] = $true; $cls[$d] = @() }
        $cls[$d] = @($cls[$d]) + $e
    }

    foreach ($n in @($script:Notes)) {
        if ($null -eq $n -or $n.done -or -not $n.due) { continue }
        $d = [string]$n.due
        if ($d -lt $todayKey) { continue }
        if (-not $days.Contains($d)) { $days[$d] = $true }
        if (-not $nts.ContainsKey($d)) { $nts[$d] = @() }
        $nts[$d] = @($nts[$d]) + $n
    }

    $keys = @($days.Keys | Sort-Object | Select-Object -First 5)
    return [pscustomobject]@{ Today = $todayKey; Keys = $keys; Classes = $cls; Notes = $nts }
}

function Save-Notes($items) {
    $arr = @()
    if ($null -ne $items) { $arr = @($items | Where-Object { $null -ne $_ }) }
    $obj = [pscustomobject]@{ items = $arr }
    $json = $obj | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($NotesPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

$Notes = Load-Notes

if ($SelfTest) {
    # 只读自检：打印「日程会显示哪几天、每天几条」，供命令行验数据与取日期的逻辑。
    $a = Get-AgendaDays
    Write-Output ('today=' + $a.Today)
    Write-Output ('notes=' + @($script:Notes).Count)
    foreach ($k in @($a.Keys)) {
        $c = 0
        if ($a.Classes.ContainsKey($k)) { $c = @($a.Classes[$k]).Count }
        $n = 0
        if ($a.Notes.ContainsKey($k)) { $n = @($a.Notes[$k]).Count }
        Write-Output ("day=$k classes=$c notes=$n")
    }
    foreach ($x in @($script:Notes)) {
        $d = 'null'
        if ($x.due) { $d = [string]$x.due }
        $dn = 'false'
        if ($x.done) { $dn = 'true' }
        Write-Output ("note due=$d done=$dn id=$($x.id)")
    }
    exit 0
}


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
    $s = @{ Left = $null; Top = $null; GpuPick = 'auto'; WeatherLat = $null; WeatherLon = $null
            PosPin = $false; PosLeft = $null; PosTop = $null }
    if (Test-Path $StatePath) {
        try {
            $o = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -ne $o.Left) { $s.Left = [double]$o.Left }
            if ($null -ne $o.Top)  { $s.Top  = [double]$o.Top }
            if ($null -ne $o.GpuPick -and "$($o.GpuPick)") { $s.GpuPick = [string]$o.GpuPick }
            if ($null -ne $o.WeatherLat -and "$($o.WeatherLat)") { $s.WeatherLat = [double]$o.WeatherLat }
            if ($null -ne $o.WeatherLon -and "$($o.WeatherLon)") { $s.WeatherLon = [double]$o.WeatherLon }
            if ("$($o.PosPin)" -eq 'True') { $s.PosPin = $true }
            if ($null -ne $o.PosLeft -and "$($o.PosLeft)") { $s.PosLeft = [double]$o.PosLeft }
            if ($null -ne $o.PosTop -and "$($o.PosTop)") { $s.PosTop = [double]$o.PosTop }
        } catch {}
    }
    return $s
}

function Save-State {
    $o = [pscustomobject]@{
        Left       = [double]$w.Left
        Top        = [double]$w.Top
        GpuPick    = [string]$script:GpuPick
        WeatherLat = $script:WeatherLat
        WeatherLon = $script:WeatherLon
        PosPin     = [bool]$script:PosPin
        PosLeft    = $script:PosLeft
        PosTop     = $script:PosTop
    }
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
          <ColumnDefinition Width="56"/>
          <ColumnDefinition Width="184"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="70"/>
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

        <Grid Grid.Column="2" VerticalAlignment="Center" Margin="0,0,4,0">
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

        <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center"
                    HorizontalAlignment="Right">
          <Border x:Name="MemoBtn" Width="20" Height="22" Background="Transparent"
                  CornerRadius="5" Cursor="Hand" Margin="0,0,1,0" ToolTip="备忘录">
            <TextBlock x:Name="MemoIcon" Text="&#xE8A5;" FontFamily="Segoe MDL2 Assets"
                       FontSize="12" Foreground="#7A7A80"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <Border x:Name="LockBtn" Width="20" Height="22" Background="Transparent"
                  CornerRadius="5" Cursor="Hand">
            <TextBlock x:Name="LockIcon" Text="&#xE785;" FontFamily="Segoe MDL2 Assets"
                       FontSize="12" Foreground="#7A7A80"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <Border x:Name="LayerBtn" Width="20" Height="22" Background="Transparent"
                  CornerRadius="5" Cursor="Hand" Margin="1,0,0,0">
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
        Width="640" Height="600"
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
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <Grid x:Name="SetHeader" Grid.Row="0" Margin="22,15,14,8" Background="Transparent">
        <TextBlock Text="设置" Foreground="#F5F5F7" FontSize="15" FontWeight="SemiBold"
                   VerticalAlignment="Center"/>
        <Button x:Name="SetCloseBtn" Content="✕" HorizontalAlignment="Right" Width="28" Height="28"
                Background="Transparent" Foreground="#8A8A90" BorderThickness="0"
                FontSize="13" Cursor="Hand" Padding="0"/>
      </Grid>

      <Grid Grid.Row="1">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="138"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="#0A0B0E" CornerRadius="0,0,0,16" Padding="10,2,10,12">
          <StackPanel>
            <Border x:Name="NavGpu" Tag="gpu" CornerRadius="8" Padding="11,8"
                    Background="Transparent" Cursor="Hand">
              <TextBlock x:Name="NavGpuText" Text="显卡监控" FontSize="12" Foreground="#8A8A90"/>
            </Border>
            <Border x:Name="NavPos" Tag="pos" CornerRadius="8" Padding="11,8" Margin="0,2,0,0"
                    Background="Transparent" Cursor="Hand">
              <TextBlock x:Name="NavPosText" Text="出现位置" FontSize="12" Foreground="#8A8A90"/>
            </Border>
            <Border x:Name="NavLoc" Tag="loc" CornerRadius="8" Padding="11,8" Margin="0,2,0,0"
                    Background="Transparent" Cursor="Hand">
              <TextBlock x:Name="NavLocText" Text="天气位置" FontSize="12" Foreground="#8A8A90"/>
            </Border>
            <Border x:Name="NavCal" Tag="cal" CornerRadius="8" Padding="11,8" Margin="0,2,0,0"
                    Background="Transparent" Cursor="Hand">
              <TextBlock x:Name="NavCalText" Text="日历订阅" FontSize="12" Foreground="#8A8A90"/>
            </Border>
            <Border x:Name="NavAbout" Tag="about" CornerRadius="8" Padding="11,8" Margin="0,2,0,0"
                    Background="Transparent" Cursor="Hand">
              <TextBlock x:Name="NavAboutText" Text="关于" FontSize="12" Foreground="#8A8A90"/>
            </Border>
          </StackPanel>
        </Border>

        <ScrollViewer Grid.Column="1" VerticalScrollBarVisibility="Auto" Margin="22,2,12,14">
          <Grid>
            <StackPanel x:Name="SetPaneGpu">
              <TextBlock Text="监控哪块显卡" Foreground="#F5F5F7" FontSize="13" FontWeight="SemiBold"
                         Margin="0,4,0,10"/>
              <WrapPanel x:Name="SetGpuRow"/>
              <TextBlock x:Name="SetGpuHint" Text="" Foreground="#6E6E73" FontSize="10.5"
                         Margin="2,8,0,0" TextWrapping="Wrap"/>
              <TextBlock Text="温度和占用每 2 秒刷新一次，切换后立即生效。"
                         Foreground="#4E4E54" FontSize="10" Margin="2,10,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <StackPanel x:Name="SetPanePos" Visibility="Collapsed">
              <TextBlock Text="面板出现位置" Foreground="#F5F5F7" FontSize="13" FontWeight="SemiBold"
                         Margin="0,4,0,6"/>
              <TextBlock Text="决定主条每次启动时出现在哪。坐标以主屏左上角为原点，左侧副屏是负数。"
                         Foreground="#6E6E73" FontSize="10.5" TextWrapping="Wrap" Margin="0,0,0,12"/>
              <WrapPanel x:Name="SetPosRow">
                <Button x:Name="SetPosPinBtn" Content="固定到当前位置" Background="#1FFFFFFF"
                        Foreground="#F5F5F7" BorderThickness="0" FontSize="11.5"
                        Padding="12,7" Margin="0,0,8,8" Cursor="Hand"/>
                <Button x:Name="SetPosFollowBtn" Content="跟随上次位置" Background="Transparent"
                        Foreground="#8A8A90" BorderThickness="0" FontSize="11.5"
                        Padding="12,7" Margin="0,0,8,8" Cursor="Hand"/>
              </WrapPanel>
              <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                <Border CornerRadius="9" Background="#1FFFFFFF" Width="104">
                  <Grid>
                    <TextBox x:Name="SetPosXInput" Background="Transparent" BorderThickness="0"
                             Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="10,8" FontSize="11.5"/>
                    <TextBlock x:Name="SetPosXPh" Text="X" IsHitTestVisible="False" Foreground="#4E4E54"
                               FontSize="11" Margin="12,0,0,0" VerticalAlignment="Center"/>
                  </Grid>
                </Border>
                <Border CornerRadius="9" Background="#1FFFFFFF" Width="104" Margin="8,0,0,0">
                  <Grid>
                    <TextBox x:Name="SetPosYInput" Background="Transparent" BorderThickness="0"
                             Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="10,8" FontSize="11.5"/>
                    <TextBlock x:Name="SetPosYPh" Text="Y" IsHitTestVisible="False" Foreground="#4E4E54"
                               FontSize="11" Margin="12,0,0,0" VerticalAlignment="Center"/>
                  </Grid>
                </Border>
                <Button x:Name="SetPosApplyBtn" Content="应用" Width="54" Height="32" Margin="8,0,0,0"
                        Background="Transparent" Foreground="#0A84FF" BorderThickness="0"
                        FontSize="11.5" Cursor="Hand" Padding="0"/>
              </StackPanel>
              <TextBlock x:Name="SetPosHint" Text="" Foreground="#6E6E73" FontSize="10"
                         Margin="2,8,0,0" TextWrapping="Wrap"/>
              <TextBlock Text="「固定」= 每次启动都回到这个点，拖动也不会记住；「跟随」= 记住你最后放的地方。"
                         Foreground="#4E4E54" FontSize="10" Margin="2,12,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <StackPanel x:Name="SetPaneLoc" Visibility="Collapsed">
              <TextBlock Text="天气位置" Foreground="#F5F5F7" FontSize="13" FontWeight="SemiBold"
                         Margin="0,4,0,6"/>
              <TextBlock Text="填这个地方的经纬度（不是窗口位置）。改完点「应用」。"
                         Foreground="#6E6E73" FontSize="10.5" TextWrapping="Wrap" Margin="0,0,0,12"/>
              <StackPanel Orientation="Horizontal">
                <Border CornerRadius="9" Background="#1FFFFFFF" Width="112">
                  <Grid>
                    <TextBox x:Name="SetLatInput" Background="Transparent" BorderThickness="0"
                             Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="10,8" FontSize="11.5"/>
                    <TextBlock x:Name="SetLatPh" Text="纬度" IsHitTestVisible="False" Foreground="#4E4E54" FontSize="11"
                               Margin="12,0,0,0" VerticalAlignment="Center"/>
                  </Grid>
                </Border>
                <Border CornerRadius="9" Background="#1FFFFFFF" Width="112" Margin="8,0,0,0">
                  <Grid>
                    <TextBox x:Name="SetLonInput" Background="Transparent" BorderThickness="0"
                             Foreground="#F5F5F7" CaretBrush="#F5F5F7" Padding="10,8" FontSize="11.5"/>
                    <TextBlock x:Name="SetLonPh" Text="经度" IsHitTestVisible="False" Foreground="#4E4E54" FontSize="11"
                               Margin="12,0,0,0" VerticalAlignment="Center"/>
                  </Grid>
                </Border>
                <Button x:Name="SetLocBtn" Content="应用" Width="54" Height="32" Margin="8,0,0,0"
                        Background="Transparent" Foreground="#0A84FF" BorderThickness="0"
                        FontSize="11.5" Cursor="Hand" Padding="0"/>
              </StackPanel>
              <TextBlock x:Name="SetLocHint" Text="" Foreground="#6E6E73" FontSize="10"
                         Margin="2,8,0,0" TextWrapping="Wrap"/>
              <TextBlock Text="两个都留空就等于用谢菲尔德；想换成别处，搜「地名 + 经纬度」填进来再点「应用」。"
                         Foreground="#4E4E54" FontSize="10" Margin="2,12,0,0" TextWrapping="Wrap"/>
            </StackPanel>

            <StackPanel x:Name="SetPaneCal" Visibility="Collapsed">
              <TextBlock Text="日历订阅" Foreground="#F5F5F7" FontSize="13" FontWeight="SemiBold"
                         Margin="0,4,0,6"/>
              <TextBlock Text="粘贴 .ics 链接（Google / Outlook / 学校教务都行）。"
                         Foreground="#6E6E73" FontSize="10.5" TextWrapping="Wrap" Margin="0,0,0,10"/>
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
              <Button x:Name="SetExportBtn" Content="导出 .ics" Background="#1FFFFFFF"
                      Foreground="#F5F5F7" BorderThickness="0" FontSize="11.5"
                      Padding="12,7" Margin="0,12,0,0" HorizontalAlignment="Left" Cursor="Hand"/>
              <TextBlock Text="把当前课程和带日期的备忘整合成一份 .ics，可以直接导进 Google / Apple / Outlook 日历。"
                         Foreground="#6E6E73" FontSize="10" Margin="2,8,0,0" TextWrapping="Wrap"/>
              <TextBlock Text="已导入的日历" Foreground="#8A8A90" FontSize="10.5" Margin="0,16,0,8"/>
              <StackPanel x:Name="SetSourceList"/>
            </StackPanel>

            <StackPanel x:Name="SetPaneAbout" Visibility="Collapsed">
              <TextBlock Text="日程球 Beacon" Foreground="#F5F5F7" FontSize="14" FontWeight="SemiBold"
                         Margin="0,4,0,4"/>
              <TextBlock Text="v1.0.11 · 日程 · 硬件监控 · 天气" Foreground="#6E6E73"
                         FontSize="10.5" TextWrapping="Wrap"/>
              <Border Margin="0,16,0,0" CornerRadius="9" Background="#12FFFFFF" Padding="12,10">
                <StackPanel>
                  <TextBlock Text="项目主页" Foreground="#8A8A90" FontSize="10"/>
                  <TextBlock x:Name="SetRepoLink" Text="github.com/T0086/beacon" Foreground="#0A84FF"
                             FontSize="11.5" Cursor="Hand" Margin="0,4,0,0"/>
                </StackPanel>
              </Border>
              <TextBlock Text="数据全部存在本地（state / notes / imported 三个 json），不上传到任何服务器。"
                         Foreground="#4E4E54" FontSize="10" TextWrapping="Wrap" Margin="2,14,0,0"/>
            </StackPanel>
          </Grid>
        </ScrollViewer>
      </Grid>
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
$MemoBtn      = $w.FindName('MemoBtn')
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
$SetGpuRow     = $sw.FindName('SetGpuRow')
$SetGpuHint    = $sw.FindName('SetGpuHint')
$SetLatInput    = $sw.FindName('SetLatInput')
$SetLonInput    = $sw.FindName('SetLonInput')
$SetLatPh       = $sw.FindName('SetLatPh')
$SetLonPh       = $sw.FindName('SetLonPh')
$SetLocBtn      = $sw.FindName('SetLocBtn')
$SetLocHint     = $sw.FindName('SetLocHint')
$SetPaneGpu     = $sw.FindName('SetPaneGpu')
$SetPanePos     = $sw.FindName('SetPanePos')
$SetPosPinBtn    = $sw.FindName('SetPosPinBtn')
$SetPosFollowBtn = $sw.FindName('SetPosFollowBtn')
$SetPosXInput    = $sw.FindName('SetPosXInput')
$SetPosYInput    = $sw.FindName('SetPosYInput')
$SetPosXPh       = $sw.FindName('SetPosXPh')
$SetPosYPh       = $sw.FindName('SetPosYPh')
$SetPosApplyBtn  = $sw.FindName('SetPosApplyBtn')
$SetPosHint      = $sw.FindName('SetPosHint')
$NavPos         = $sw.FindName('NavPos')
$NavPosText     = $sw.FindName('NavPosText')
$SetPaneLoc     = $sw.FindName('SetPaneLoc')
$SetPaneCal     = $sw.FindName('SetPaneCal')
$SetPaneAbout   = $sw.FindName('SetPaneAbout')
$NavGpu         = $sw.FindName('NavGpu')
$NavLoc         = $sw.FindName('NavLoc')
$NavCal         = $sw.FindName('NavCal')
$NavAbout       = $sw.FindName('NavAbout')
$NavGpuText     = $sw.FindName('NavGpuText')
$NavLocText     = $sw.FindName('NavLocText')
$NavCalText     = $sw.FindName('NavCalText')
$NavAboutText   = $sw.FindName('NavAboutText')
$SetRepoLink    = $sw.FindName('SetRepoLink')
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
# 安全启动：上次如果没干净退出，lock 还留着，就别信「上次位置」——
# 崩溃/断电时存下来的坐标经常是半截的，照着恢复会让条子跑到屏幕外或者叠在任务栏上。
$LockPath = Join-Path $Base 'ctl\running.lock'
$safeMode = Test-Path $LockPath
try { New-Item -ItemType Directory -Force -Path (Split-Path $LockPath) | Out-Null } catch {}
try { Set-Content -Path $LockPath -Value (Get-Date -Format 'o') -Encoding UTF8 } catch {}
$wa = [System.Windows.SystemParameters]::WorkArea
$state = Load-State
$script:GpuPick    = [string]$state.GpuPick
$script:WeatherLat = $state.WeatherLat
$script:WeatherLon = $state.WeatherLon
$script:PosPin     = [bool]$state.PosPin
$script:PosLeft    = $state.PosLeft
$script:PosTop     = $state.PosTop
# 固定位置优先：勾了「固定」就用那组坐标，不再跟随上次拖动
if ($script:PosPin -and $null -ne $script:PosLeft -and $null -ne $script:PosTop) {
    $w.Left = $script:PosLeft
    $w.Top  = $script:PosTop
} elseif (-not $safeMode -and $null -ne $state.Left -and $null -ne $state.Top) {
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

function Get-LhmGpuList {
    $out = @()
    if (-not $script:LhmComputer) { return $out }
    try {
        foreach ($h in $script:LhmComputer.Hardware) {
            if ("$($h.HardwareType)" -notmatch 'Gpu') { continue }
            $out += "$($h.Name)"
        }
    } catch {}
    return $out
}

function Get-LhmGpuLoad {
    if (-not $script:LhmComputer) { return $null }
    try {
        $pick = if ($script:GpuPick) { [string]$script:GpuPick } else { 'auto' }
        $entries = @()
        foreach ($h in $script:LhmComputer.Hardware) {
            $type = "$($h.HardwareType)"
            if ($type -notmatch 'Gpu') { continue }
            # honour the user's explicit adapter choice from Settings
            if ($pick -ne 'auto' -and "$($h.Name)" -ne $pick) { continue }
            $h.Update()
            $load = $null
            foreach ($s in $h.Sensors) {
                if ("$($s.SensorType)" -eq 'Load' -and "$($s.Name)" -eq 'GPU Core' -and $null -ne $s.Value) {
                    $load = [double]$s.Value
                    break
                }
            }
            if ($null -ne $load) {
                $entries += [pscustomobject]@{ Type = $type; Name = "$($h.Name)"; Load = $load }
            }
        }
        if ($entries.Count -eq 0) { return $null }

        # 1) NVIDIA adapters are always discrete - prefer them
        $nv = $entries | Where-Object { $_.Type -eq 'GpuNvidia' } | Select-Object -First 1
        if ($nv) { return $nv.Load }

        # 2) discrete AMD / Intel cards usually carry "RX " / "Radeon Pro" / "Arc "
        $disc = $entries | Where-Object { $_.Name -match 'RX |Radeon Pro|Arc ' } | Select-Object -First 1
        if ($disc) { return $disc.Load }

        # 3) single-adapter or integrated-only machine: take the busiest one
        return ($entries | Sort-Object -Property Load -Descending | Select-Object -First 1).Load
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
        $lat = if ($null -ne $script:WeatherLat) { [double]$script:WeatherLat } else { 53.3811 }
        $lon = if ($null -ne $script:WeatherLon) { [double]$script:WeatherLon } else { -1.4701 }
        $url = "https://api.open-meteo.com/v1/forecast?latitude=$lat&longitude=$lon&current=temperature_2m,weather_code"
        $r = Invoke-RestMethod -Uri $url -TimeoutSec 8
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
        $t.Foreground = Brush '#A8A8B0'
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

# ---- left nav: one pane at a time, nothing stacked on one screen
$script:SetTab = 'gpu'
function Select-Pane([string]$key) {
    $script:SetTab = $key
    $rows = @(
        @{ k = 'gpu';   nav = $NavGpu;   txt = $NavGpuText;   pane = $SetPaneGpu },
        @{ k = 'pos';   nav = $NavPos;   txt = $NavPosText;   pane = $SetPanePos },
        @{ k = 'loc';   nav = $NavLoc;   txt = $NavLocText;   pane = $SetPaneLoc },
        @{ k = 'cal';   nav = $NavCal;   txt = $NavCalText;   pane = $SetPaneCal },
        @{ k = 'about'; nav = $NavAbout; txt = $NavAboutText; pane = $SetPaneAbout }
    )
    foreach ($r in $rows) {
        if ($r.k -eq $key) {
            $r.nav.Background = Brush '#1FFFFFFF'
            $r.txt.Foreground = Brush '#F5F5F7'
            $r.txt.FontWeight = 'SemiBold'
            $r.pane.Visibility = 'Visible'
        } else {
            $r.nav.Background = Brush 'Transparent'
            $r.txt.Foreground = Brush '#8A8A90'
            $r.txt.FontWeight = 'Normal'
            $r.pane.Visibility = 'Collapsed'
        }
    }
}

$NavGpu.Add_MouseLeftButtonUp({ Select-Pane 'gpu' })
$NavPos.Add_MouseLeftButtonUp({ Select-Pane 'pos' })
$NavLoc.Add_MouseLeftButtonUp({ Select-Pane 'loc' })
$NavCal.Add_MouseLeftButtonUp({ Select-Pane 'cal' })
$NavAbout.Add_MouseLeftButtonUp({ Select-Pane 'about' })

$SetRepoLink.Add_MouseLeftButtonUp({
    try { [System.Diagnostics.Process]::Start('https://github.com/T0086/beacon') } catch {}
})

$SetUrlInput.Add_TextChanged({
    if ($SetUrlInput.Text) { $SetUrlHint.Visibility = 'Collapsed' } else { $SetUrlHint.Visibility = 'Visible' }
})

$SetImportBtn.Add_Click({ Import-IcsSource })
$SetExportBtn  = $sw.FindName('SetExportBtn')
$SetExportBtn.Add_Click({ Export-Ics })

# apply the weather location (lat / lon)
$SetLocBtn.Add_Click({
    $laTxt = ([string]$SetLatInput.Text).Trim()
    $loTxt = ([string]$SetLonInput.Text).Trim()
    if (-not $laTxt -and -not $loTxt) {
        # both empty = clear the custom location, fall back to the built-in default
        $script:WeatherLat = $null
        $script:WeatherLon = $null
        $script:WeatherAt  = [datetime]::MinValue
        Save-State
        Update-Weather
        $SetLocHint.Text = '已清空，天气按谢菲尔德查'
        return
    }
    try {
        $la = [double]::Parse($laTxt, [System.Globalization.CultureInfo]::InvariantCulture)
        $lo = [double]::Parse($loTxt, [System.Globalization.CultureInfo]::InvariantCulture)
        if ($la -lt -90 -or $la -gt 90) { throw 'lat' }
        if ($lo -lt -180 -or $lo -gt 180) { throw 'lon' }
        $script:WeatherLat = $la
        $script:WeatherLon = $lo
        $script:WeatherAt  = [datetime]::MinValue
        Save-State
        Update-Weather
        $SetLocHint.Text = "已保存 $la, $lo"
    } catch {
        $SetLocHint.Text = '请输入有效数字（纬度 -90~90，经度 -180~180）'
    }
})

# watermark: hide the 纬度 / 经度 hint as soon as the box has text in it
function Sync-LocPlaceholder {
    $SetLatPh.Visibility = if ("$($SetLatInput.Text)") { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    $SetLonPh.Visibility = if ("$($SetLonInput.Text)") { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
}
$SetLatInput.Add_TextChanged({ Sync-LocPlaceholder })
$SetLonInput.Add_TextChanged({ Sync-LocPlaceholder })

# ---- 出现位置：固定 or 跟随上次
function Sync-PosPlaceholder {
    $SetPosXPh.Visibility = if ("$($SetPosXInput.Text)") { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    $SetPosYPh.Visibility = if ("$($SetPosYInput.Text)") { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
}
$SetPosXInput.Add_TextChanged({ Sync-PosPlaceholder })
$SetPosYInput.Add_TextChanged({ Sync-PosPlaceholder })

function Set-PosPin([bool]$pin, [double]$left, [double]$top) {
    $script:PosPin = $pin
    if ($pin) {
        $script:PosLeft = $left
        $script:PosTop  = $top
        $w.Left = $left
        $w.Top  = $top
    }
    Save-State
    Render-Settings
    Select-Pane 'pos'
}

$SetPosPinBtn.Add_Click({
    Set-PosPin $true $w.Left $w.Top
    $SetPosHint.Text = '已固定到 ' + [math]::Round($w.Left) + ', ' + [math]::Round($w.Top)
})
$SetPosFollowBtn.Add_Click({
    Set-PosPin $false 0 0
    $SetPosHint.Text = '已改为跟随上次位置'
})
$SetPosApplyBtn.Add_Click({
    try {
        $x = [double]::Parse(([string]$SetPosXInput.Text).Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
        $y = [double]::Parse(([string]$SetPosYInput.Text).Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
        Set-PosPin $true $x $y
        $SetPosHint.Text = "已固定到 $x, $y"
    } catch {
        $SetPosHint.Text = '请输入数字（左侧副屏是负数，例如 -1912）'
    }
})

$SettingsBtn.Add_Click({
    if (-not $sw.IsVisible) {
        $setLeft = $w.Left - 640 - 12
        if ($setLeft -lt $wa.Left + 8) { $setLeft = $wa.Left + 8 }
        $sw.Left = $setLeft
        $sw.Top  = [math]::Max($wa.Top + 8, [math]::Min($w.Top, $wa.Top + $wa.Height - 620))
        $sw.Show()
    } else {
        $sw.Activate()
    }
    Render-Settings
    Select-Pane $script:SetTab
})

# ---------------------------------------------------------------- lock / layer
$script:Locked = $false
$script:OnTop  = $true

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class OrbWin32 {
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
  [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
  [DllImport("user32.dll")] public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
"@ -ErrorAction SilentlyContinue

# WS_EX_TOPMOST on the real window versus $script:OnTop are two different things and they
# drift: a hidden-then-shown window, a fullscreen app taking the top, a game changing mode.
# Toggling the cached boolean then costs two clicks - the first one clears a flag that was
# already effectively off, and only the second one actually raises the window. So never
# trust the variable: ask the window, then force the z-order with SetWindowPos.
function Get-OnTopActual {
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
        if ($hwnd -eq [IntPtr]::Zero) { return $script:OnTop }
        $ex = [int][OrbWin32]::GetWindowLong($hwnd, -20)
        return (($ex -band 0x00000008) -ne 0)
    } catch { return $script:OnTop }
}

function Set-OnTop([bool]$on) {
    $script:OnTop = $on
    try { $w.Topmost = $on } catch {}
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
        if ($hwnd -ne [IntPtr]::Zero) {
            # HWND_TOPMOST = -1, HWND_NOTOPMOST = -2, SWP_NOMOVE|SWP_NOSIZE = 0x0003
            $after = [IntPtr]$(if ($on) { -1 } else { -2 })
            [void][OrbWin32]::SetWindowPos($hwnd, $after, 0, 0, 0, 0, 0x0003)
        }
    } catch {}
    try { Update-LayerUI } catch {}
}

# Click-through: only WS_EX_TRANSPARENT really forwards clicks to *other processes*.
# WM_NCHITTEST / HTTRANSPARENT only skips between windows of the same thread, which is
# why clicking through to the browser did nothing. To keep the padlock usable we park a
# tiny invisible always-on-top window exactly over it; that one eats the unlock click
# while everything else on the bar lets the mouse reach the app underneath.
$script:ClickThrough = $false

function Set-ClickThrough([bool]$on) {
    if ($script:ClickThrough -eq $on) { return }
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
        $ex = [int][OrbWin32]::GetWindowLong($hwnd, -20)
        if ($on) { $ex = $ex -bor 0x00000020 } else { $ex = $ex -band (-bnot 0x00000020) }
        [void][OrbWin32]::SetWindowLong($hwnd, -20, $ex)
        $script:ClickThrough = $on
    } catch {}
}

$script:UnlockWin = $null

# WS_EX_NOACTIVATE (0x08000000): the bar never takes activation, so clicking it
# while a game is in the foreground no longer pulls focus out of the game.
function Set-NoActivate {
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
        $ex = [int][OrbWin32]::GetWindowLong($hwnd, -20)
        if (($ex -band 0x08000000) -eq 0) {
            [void][OrbWin32]::SetWindowLong($hwnd, -20, ($ex -bor 0x08000000))
        }
    } catch {}
}

$w.Add_SourceInitialized({ Set-NoActivate; if ($script:OnTop) { Set-OnTop $true } })

# 2026-11-21：NOACTIVATE 让这个窗口永远拿不到键盘焦点，备忘输入框于是打不进字。
# 点输入框时临时允许激活并抢一次焦点，窗口失焦后再恢复 NOACTIVATE（游戏里仍然不抢焦点）。
function Set-Activatable([bool]$on) {
    try {
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper($w)).Handle
        $ex = [int][OrbWin32]::GetWindowLong($hwnd, -20)
        if ($on) { $ex = $ex -band (-bnot 0x08000000) } else { $ex = $ex -bor 0x08000000 }
        [void][OrbWin32]::SetWindowLong($hwnd, -20, $ex)
    } catch {}
}

function Show-UnlockPatch {
    try {
        if (-not $script:UnlockWin) {
            $uw = New-Object System.Windows.Window
            $uw.WindowStyle       = 'None'
            $uw.AllowsTransparency = $true
            $uw.Background        = Brush '#01161920'   # alpha 1/255: invisible yet hit-testable
            $uw.ShowInTaskbar     = $false
            $uw.Topmost           = $true
            $uw.ResizeMode        = 'NoResize'
            $uw.ShowActivated     = $false
            $uw.Cursor            = 'Hand'
            $uw.Add_MouseLeftButtonUp({
                param($s, $ev)
                $ev.Handled = $true
                # 2026-11-21 KJ：锁定时左半边是「备忘录」（只开详细页，不解锁），右半边才是解锁
                $x = 0
                try { $x = $ev.GetPosition($script:UnlockWin).X } catch {}
                if ($script:PatchMemoW -gt 0 -and $x -lt $script:PatchMemoW) {
                    Expand-Orb
                    $script:PeekOpen = $true
                    Set-ClickThrough $false
                } else {
                    $script:Locked = $false
                    Update-LockUI
                }
            })
            $script:UnlockWin = $uw
        }
        $脚本 = $script:UnlockWin
        if ($MemoBtn -and $MemoBtn.ActualWidth -gt 0 -and $LockBtn.ActualWidth -gt 0) {
            $p1 = $MemoBtn.TranslatePoint((New-Object System.Windows.Point(0, 0)), $w)
            $p2 = $LockBtn.TranslatePoint((New-Object System.Windows.Point(0, 0)), $w)
            $script:PatchMemoW = ([double]$MemoBtn.ActualWidth) + 2
            $脚本.Width  = [double]($p2.X + $LockBtn.ActualWidth - $p1.X) + 10
            $脚本.Height = [double]([Math]::Max($LockBtn.ActualHeight, $MemoBtn.ActualHeight)) + 10
            $脚本.Left   = $w.Left + $p1.X - 5
            $脚本.Top    = $w.Top  + $p1.Y - 5
        } elseif ($LockBtn.ActualWidth -gt 0) {
            $p = $LockBtn.TranslatePoint((New-Object System.Windows.Point(0, 0)), $w)
            $script:PatchMemoW = 0
            $脚本.Width  = [double]$LockBtn.ActualWidth + 10
            $脚本.Height = [double]$LockBtn.ActualHeight + 10
            $脚本.Left   = $w.Left + $p.X - 5
            $脚本.Top    = $w.Top  + $p.Y - 5
        }
        if (-not $脚本.IsVisible) { $脚本.Show() }
    } catch {}
}

function Hide-UnlockPatch {
    try { if ($script:UnlockWin) { $script:UnlockWin.Hide() } } catch {}
}

function Update-LockUI {
    if ($script:Locked) {
        $LockIcon.Text = [string][char]0xE72E
        $LockIcon.Foreground = Brush '#FF9F0A'
        # locked: fade the bar down so underlying content stays readable
        try { $OrbView.Background = Brush '#8C161920' } catch {}
        Show-UnlockPatch
        Set-ClickThrough $true
    } else {
        $LockIcon.Text = [string][char]0xE785
        $LockIcon.Foreground = Brush '#7A7A80'
        try { $OrbView.Background = Brush '#F0181B21' } catch {}
        Set-ClickThrough $false
        Hide-UnlockPatch
    }
}

function Update-LayerUI {
    $LayerIcon.Foreground = Brush $(if ($script:OnTop) { '#0A84FF' } else { '#5E5E64' })
}

$LockBtn.Add_MouseLeftButtonDown({
    param($s, $ev)
    $ev.Handled = $true
    # Handled=true stops the parent from starting a drag, but it also skips ButtonBase's own
    # capture - and without capture a click that moves 2px loses its ButtonUp. Hold it here.
    try { [void]$LockBtn.CaptureMouse() } catch {}
})
$LayerBtn.Add_MouseLeftButtonDown({
    param($s, $ev)
    $ev.Handled = $true
    try { [void]$LayerBtn.CaptureMouse() } catch {}
})

$LockBtn.Add_MouseLeftButtonUp({
    param($s, $ev)
    $ev.Handled = $true
    try { [void]$LockBtn.ReleaseMouseCapture() } catch {}
    $script:Locked = -not $script:Locked
    Update-LockUI
})

$MemoBtn.Add_MouseLeftButtonDown({
    param($s, $ev)
    $ev.Handled = $true
    try { [void]$MemoBtn.CaptureMouse() } catch {}
})
$MemoBtn.Add_MouseLeftButtonUp({
    param($s, $ev)
    $ev.Handled = $true
    try { [void]$MemoBtn.ReleaseMouseCapture() } catch {}
    Expand-Orb
})
$LayerBtn.Add_MouseLeftButtonUp({
    param($s, $ev)
    $ev.Handled = $true
    try { [void]$LayerBtn.ReleaseMouseCapture() } catch {}
    Set-OnTop (-not (Get-OnTopActual))
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

function Dump-AgendaUi {
    # walk what is actually painted in ClassList and write it out line by line.
    # used from ctl: "dump" -> ctl\ui-dump.txt, so a command line can check the panel without a screen.
    try {
        $lines = @()
        foreach ($c in $ClassList.Children) {
            $txts = @()
            $queue = @($c)
            $guard = 0
            while ($queue.Count -gt 0 -and $guard -lt 4000) {
                $guard++
                $cur = $queue[0]
                $queue = @($queue | Select-Object -Skip 1)
                if ($null -eq $cur) { continue }
                if ($cur -is [System.Windows.Controls.TextBlock]) { $txts += [string]$cur.Text; continue }
                if ($cur -is [System.Windows.Controls.Border] -and $cur.Child) { $queue += $cur.Child; continue }
                if ($cur -is [System.Windows.Controls.StackPanel]) { foreach ($ch in $cur.Children) { $queue += $ch }; continue }
                if ($cur -is [System.Windows.Controls.Grid]) { foreach ($ch in $cur.Children) { $queue += $ch }; continue }
            }
            $lines += ($txts -join ' | ')
        }
        $outPath = Join-Path $CtlDir 'ui-dump.txt'
        [System.IO.File]::WriteAllText($outPath, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    } catch { Write-Err 'dump' }
}
function Render-Classes {
    $ClassList.Children.Clear()
    $a = Get-AgendaDays
    $keys = @($a.Keys)

    if ($keys.Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '这一学期没有更多课程了'
        $t.Foreground = Brush '#A8A8B0'
        $t.FontSize = 12
        $t.Margin = '6,24,0,0'
        $ClassList.Children.Add($t) | Out-Null
        return
    }

    foreach ($k in $keys) {
        Add-DayHeader $k
        if ($a.Classes.ContainsKey($k)) {
            foreach ($e in @($a.Classes[$k])) { Add-ClassRow $e }
        }
        if ($a.Notes.ContainsKey($k)) {
            foreach ($n in @($a.Notes[$k])) { Add-DeadlineRow $n }
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
        $t.Foreground = Brush '#A8A8B0'
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
        $btn.Foreground = Brush '#F2F2F5'
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
        $t.Foreground = Brush '#A8A8B0'
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
    # ---- GPU selector
    $SetGpuRow.Children.Clear()
    $mkGpuBtn = {
        param($label, $value)
        $b = New-Object System.Windows.Controls.Border
        $b.CornerRadius = '7'
        $b.Padding = '10,5,10,5'
        $b.Margin = '0,0,6,0'
        $b.Cursor = 'Hand'
        $b.Tag = $value
        $b.BorderThickness = '1'
        if ("$($script:GpuPick)" -eq "$value") {
            $b.Background = Brush '#33FF9F0A'
            $b.BorderBrush = Brush '#FF9F0A'
        } else {
            $b.Background = Brush '#14FFFFFF'
            $b.BorderBrush = Brush 'Transparent'
        }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $label
        $tb.FontSize = 10.5
        $tb.TextTrimming = 'CharacterEllipsis'
        $tb.MaxWidth = 150
        $tb.Foreground = Brush $(if ("$($script:GpuPick)" -eq "$value") { '#FFBF4A' } else { '#B0B0B6' })
        $b.Child = $tb
        $b.Add_MouseLeftButtonUp({
            param($sender, $ev)
            $script:GpuPick = [string]$sender.Tag
            Save-State
            Render-Settings
        })
        $SetGpuRow.Children.Add($b) | Out-Null
    }

    & $mkGpuBtn '自动' 'auto'
    $gpuNames = Get-LhmGpuList
    foreach ($gn in $gpuNames) { & $mkGpuBtn $gn $gn }

    if (-not $script:LhmComputer) {
        $SetGpuHint.Text = '硬件监控库未加载，无法读取显卡'
    } elseif (@($gpuNames).Count -le 1) {
        $SetGpuHint.Text = '只检测到一块显卡'
    } else {
        $SetGpuHint.Text = '自动 = 优先独显，没有独显就用核显'
    }

    # ---- appear position
    if ($script:PosPin -and $null -ne $script:PosLeft -and $null -ne $script:PosTop) {
        $SetPosXInput.Text = [string]$script:PosLeft
        $SetPosYInput.Text = [string]$script:PosTop
        $SetPosHint.Text = '已固定 · 每次启动都回到这里'
    } else {
        $SetPosXInput.Text = ''
        $SetPosYInput.Text = ''
        $SetPosHint.Text = '跟随上次位置（现在 ' + [math]::Round($w.Left) + ', ' + [math]::Round($w.Top) + '）'
    }
    Sync-PosPlaceholder

    # ---- weather location fields
    if ($null -ne $script:WeatherLat -and $null -ne $script:WeatherLon) {
        $SetLatInput.Text = [string]$script:WeatherLat
        $SetLonInput.Text = [string]$script:WeatherLon
        $SetLocHint.Text = '当前坐标'
    } else {
        $SetLatInput.Text = ''
        $SetLonInput.Text = ''
        $SetLocHint.Text = '还没设置，天气按谢菲尔德（53.3811, -1.4701）查'
    }
    Sync-LocPlaceholder

    $SetSourceList.Children.Clear()
    if (@($script:Sources).Count -eq 0) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = '还没有导入任何日历'
        $t.Foreground = Brush '#A8A8B0'
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

# ---------------------------------------------------------------- 导出 ics
# 把现在手里的东西整合成一份标准 .ics：课程事件（带时间的）+ 带截止日的备忘（全天）。
# 时间写成「不带时区的本地时间」——个人日历按本机时区理解，不必为 Europe/London
# 的夏令时维护一整段 VTIMEZONE；换到别的时区日历里也仍然是你当时写下的那个钟点。
function Escape-IcsText([string]$s) {
    if (-not $s) { return '' }
    $t = $s -replace '\\', '\\\\'
    $t = $t -replace ';', '\;'
    $t = $t -replace ',', '\,'
    $t = $t -replace "`r`n", '\n'
    $t = $t -replace "`n", '\n'
    return $t
}

function New-IcsUid([string]$seed) {
    try {
        $sha = [System.Security.Cryptography.SHA1]::Create()
        $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($seed))
        return ([System.BitConverter]::ToString($hash) -replace '-', '').Substring(0, 24).ToLower() + '@beacon'
    } catch {
        return ([guid]::NewGuid().ToString('N').Substring(0, 24)) + '@beacon'
    }
}

function Build-Ics {
    $nl = "`r`n"
    $lines = New-Object System.Collections.ArrayList
    [void]$lines.Add('BEGIN:VCALENDAR')
    [void]$lines.Add('VERSION:2.0')
    [void]$lines.Add('PRODID:-//beacon//beacon//ZH')
    [void]$lines.Add('CALSCALE:GREGORIAN')
    [void]$lines.Add('METHOD:PUBLISH')
    [void]$lines.Add('X-WR-CALNAME:beacon')
    $stamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $count = 0

    foreach ($e in @($script:Events)) {
        if (-not $e -or -not $e.start) { continue }
        $st = [string]$e.start
        if ($st -notmatch '^\d{4}-\d{2}-\d{2}T') { continue }
        [void]$lines.Add('BEGIN:VEVENT')
        [void]$lines.Add('UID:' + (New-IcsUid ($st + '|' + [string]$e.title)))
        [void]$lines.Add('DTSTAMP:' + $stamp)
        [void]$lines.Add('DTSTART:' + ($st -replace '[-:]', ''))
        if ($e.end) {
            $en = [string]$e.end
            if ($en -match '^\d{4}-\d{2}-\d{2}T') { [void]$lines.Add('DTEND:' + ($en -replace '[-:]', '')) }
        }
        [void]$lines.Add('SUMMARY:' + (Escape-IcsText ([string]$e.title)))
        if ($e.location) { [void]$lines.Add('LOCATION:' + (Escape-IcsText ([string]$e.location))) }
        [void]$lines.Add('END:VEVENT')
        $count++
    }

    foreach ($n in @($script:Notes)) {
        if (-not $n -or $n.done -or -not $n.due) { continue }
        $due = [string]$n.due
        if ($due -notmatch '^\d{4}-\d{2}-\d{2}$') { continue }
        $noteText = [string]$n.text

        # 时间规则（KJ 定的）：备注里写了 HH:mm 就用它；没写就落到晚上 22:00。
        # 时长：写了「X 小时 / X分钟」按它算，没写就一小时。
        $hh = 22; $mm = 0
        $tm = [regex]::Match($noteText, '(?<!\d)([01]?\d|2[0-3])[:：]([0-5]\d)(?!\d)')
        if ($tm.Success) {
            $hh = [int]$tm.Groups[1].Value
            $mm = [int]$tm.Groups[2].Value
        }
        $mins = 60
        $dm = [regex]::Match($noteText, '(\d+(?:\.\d+)?)\s*(小时|h|hr|hour)')
        if ($dm.Success) {
            $mins = [int][math]::Round([double]$dm.Groups[1].Value * 60)
        } else {
            $dm2 = [regex]::Match($noteText, '(\d+)\s*(分钟|min)')
            if ($dm2.Success) { $mins = [int]$dm2.Groups[1].Value }
        }
        if ($mins -lt 15) { $mins = 15 }
        if ($mins -gt 12 * 60) { $mins = 12 * 60 }

        $startAt = $null
        try { $startAt = [datetime]::ParseExact($due, 'yyyy-MM-dd', $null).AddHours($hh).AddMinutes($mm) } catch { continue }
        $endAt = $startAt.AddMinutes($mins)

        [void]$lines.Add('BEGIN:VEVENT')
        [void]$lines.Add('UID:' + (New-IcsUid ('note|' + [string]$n.id + '|' + $due)))
        [void]$lines.Add('DTSTAMP:' + $stamp)
        [void]$lines.Add('DTSTART:' + $startAt.ToString('yyyyMMddTHHmmss'))
        [void]$lines.Add('DTEND:' + $endAt.ToString('yyyyMMddTHHmmss'))
        [void]$lines.Add('SUMMARY:' + (Escape-IcsText $noteText))
        [void]$lines.Add('DESCRIPTION:' + (Escape-IcsText ('备忘 · 来自 beacon · ' + $mins + ' 分钟')))
        [void]$lines.Add('TRANSP:TRANSPARENT')
        [void]$lines.Add('END:VEVENT')
        $count++
    }

    [void]$lines.Add('END:VCALENDAR')
    return [pscustomobject]@{ text = (($lines -join $nl) + $nl); count = $count }
}

function Export-Ics {
    try {
        $built = Build-Ics
        if ($built.count -eq 0) {
            $SetStatus.Text = '现在没有可导出的事件'
            $SetStatus.Foreground = Brush '#FF9F0A'
            return
        }
        $expDir = Join-Path $Base 'export'
        if (-not (Test-Path $expDir)) { New-Item -ItemType Directory -Path $expDir -Force | Out-Null }

        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'iCalendar (*.ics)|*.ics'
        $dlg.DefaultExt = 'ics'
        $dlg.FileName = 'beacon-' + (Get-Date -Format 'yyyyMMdd') + '.ics'
        $dlg.InitialDirectory = $expDir
        if ($dlg.ShowDialog() -ne $true) { return }

        [System.IO.File]::WriteAllText($dlg.FileName, $built.text, (New-Object System.Text.UTF8Encoding($false)))
        $SetStatus.Text = '已导出 ' + $built.count + ' 条 → ' + [System.IO.Path]::GetFileName($dlg.FileName)
        $SetStatus.Foreground = Brush '#30D158'
    } catch {
        $SetStatus.Text = '导出失败：' + $_.Exception.Message
        $SetStatus.Foreground = Brush '#FF453A'
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
        Render-Classes
    } catch {
        Write-Err 'add-note'
    }
}

$NoteInput.Add_PreviewMouseLeftButtonDown({
    Set-Activatable $true
    try { $w.Activate() } catch {}
})
$w.Add_Deactivated({ Set-Activatable $false })

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
    # 置顶 drifts while the bar sits collapsed; force the z-order before anything else
    if ($script:OnTop) { Set-OnTop $true }
    if ($script:IsOpen) { return }

    try {
        Reload-Notes
        Render-Classes
        Render-Notes
    } catch {
        Write-Err 'render'
        return
    }

    $script:IsOpen = $true

    # 收起时的条是锚：面板怎么摆都从它算，收起时原样还回去。
    $script:BarLeft = $w.Left
    $script:BarTop  = $w.Top

    # 面板贴哪一边，看条落在屏幕的哪个半边：靠左贴左，靠右贴右；
    # 上下同理，条在下半屏时面板往上长。原来是永远贴右、永远往下，条在左边时
    # 面板离屏幕边还差 90px，看着很别扭。
    $scr = $null
    try {
        $pt = New-Object System.Drawing.Point([int]($script:BarLeft + 235), [int]($script:BarTop + 38))
        $scr = [System.Windows.Forms.Screen]::FromPoint($pt)
    } catch {}
    if (-not $scr) { $scr = [System.Windows.Forms.Screen]::PrimaryScreen }
    $b = $scr.Bounds

    $newLeft = $(if (($script:BarLeft + 235) -lt ($b.Left + $b.Width / 2)) { $script:BarLeft } else { $script:BarLeft + 470 - 380 })
    $newTop  = $(if (($script:BarTop + 38) -lt ($b.Top + $b.Height / 2)) { $script:BarTop } else { $script:BarTop + 76 - 560 })

    # 再钳一次，别把面板推出这块屏
    if ($newLeft -lt $b.Left) { $newLeft = $b.Left }
    if (($newLeft + 380) -gt $b.Right) { $newLeft = $b.Right - 380 }
    if ($newTop -lt $b.Top) { $newTop = $b.Top }
    if (($newTop + 560) -gt $b.Bottom) { $newTop = $b.Bottom - 560 }

    $w.Width  = 380
    $w.Height = 560
    $w.Left = $newLeft
    $w.Top  = $newTop

    $OrbView.Visibility = 'Collapsed'
    $PanelView.Visibility = 'Visible'

    # the month calendar is the default view whenever the panel opens
    try { Switch-Tab 'month' } catch { Write-Err 'tab' }
}

function Collapse-Orb {
    if (-not $script:IsOpen) { return }
    $script:IsOpen = $false

    # 回到展开前的那个位置，不再假设"右下角锚定"
    $w.Width  = 470
    $w.Height = 76
    $w.Left = $script:BarLeft
    $w.Top  = $script:BarTop

    $PanelView.Visibility = 'Collapsed'
    $OrbView.Visibility = 'Visible'
    # 锁定状态下看过详细页，收起后要恢复穿透
    if ($script:PeekOpen) { $script:PeekOpen = $false; if ($script:Locked) { Set-ClickThrough $true } }
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
# when locked the window is click-through, so the tray is the only way back
$miUnlock = New-Object System.Windows.Forms.ToolStripMenuItem '解除锁定'
$miUnlock.Add_Click({
    if ($script:Locked) {
        $script:Locked = $false
        Update-LockUI
    }
})
$miQuit = New-Object System.Windows.Forms.ToolStripMenuItem '退出'
$miQuit.Add_Click({
    Save-State
    $ni.Visible = $false
    $w.Close()
})
$menu.Items.Add($miShow) | Out-Null
$menu.Items.Add($miUnlock) | Out-Null
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
        $pm = $MemoBtn.TranslatePoint($p0, $w)
        $pg = $LayerBtn.TranslatePoint($p0, $w)
        $need = $pg.X + $LayerBtn.ActualWidth
        $msg = "window=$($w.ActualWidth)x$($w.ActualHeight)`r`n" +
               "orb=$($OrbView.ActualWidth)x$($OrbView.ActualHeight) at $($po.X),$($po.Y)`r`n" +
               "memo=$($MemoBtn.ActualWidth)x$($MemoBtn.ActualHeight) at $($pm.X),$($pm.Y)`r`n" +
               "lock=$($LockBtn.ActualWidth)x$($LockBtn.ActualHeight) at $($pl.X),$($pl.Y)`r`n" +
               "layer=$($LayerBtn.ActualWidth)x$($LayerBtn.ActualHeight) at $($pg.X),$($pg.Y)`r`n" +
               "layerRight=$([Math]::Round($need,1)) / windowRight=$($w.ActualWidth) -> " +
               $(if ($need -le $w.ActualWidth) { 'OK 完整可见' } else { "被裁 $([Math]::Round($need - $w.ActualWidth,1))px" })
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

# ---------------------------------------------------------------- graceful shutdown
# The tray icon only goes away if the process disposes it itself; Stop-Process leaves a
# ghost. Every shutdown path therefore funnels through this one function.
function Invoke-BeaconStop {
    try { Save-State } catch {}
    # 正常退出：清掉 lock，下次启动才不算「意外退出」
    try { Remove-Item (Join-Path $Base 'ctl\running.lock') -Force -ErrorAction SilentlyContinue } catch {}
    try { $ni.Visible = $false; $ni.Dispose() } catch {}
    try { $script:OrbMutex.ReleaseMutex() } catch {}
    try { $w.Close() } catch {}
    # closing the window alone does not always end the process, and a lingering
    # process is what leaves the ghost tray icon behind
    try { [System.Windows.Application]::Current.Shutdown() } catch {}
    [Environment]::Exit(0)
}

# graceful shutdown channel: writing quit.flag lets a restart dispose the tray icon
# properly instead of leaving a ghost behind (Stop-Process can't clean up NotifyIcon)
$quitWatcher = New-Object System.Windows.Threading.DispatcherTimer
$quitWatcher.Interval = [TimeSpan]::FromSeconds(2)
$quitWatcher.Add_Tick({
    $qf = Join-Path $Base 'quit.flag'
    if (Test-Path $qf) {
        Remove-Item $qf -Force -ErrorAction SilentlyContinue
        Invoke-BeaconStop
    }
})
$quitWatcher.Start()

# ---------------------------------------------------------------- control channel
# This process is meant to outlive the shells that talk to it, so nobody holds a handle
# to it. Both sides use files under ctl\ instead:
#   cmd.txt      verbs to run, one per line; consumed and deleted by beacon
#   status.json  heartbeat rewritten every 2s (pid, startedAt, uptimeSec, expanded, locked)
#   result.txt   outcome of the last verb, tab separated
$script:StartedAt = Get-Date

function Write-CtlFile([string]$Name, [string]$Text) {
    $path = Join-Path $CtlDir $Name
    $tmp  = "$path.tmp"
    [System.IO.File]::WriteAllText($tmp, $Text, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $path -Force
}

function Write-CtlResult([string]$Verb, [string]$Status) {
    try {
        $at = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Write-CtlFile 'result.txt' "$at`t$Verb`t$Status"
    } catch {}
}

function Write-BeaconHeartbeat {
    try {
        $up = [int]((Get-Date) - $script:StartedAt).TotalSeconds
        $exp = 'false'
        if ($script:IsOpen) { $exp = 'true' }
        $lck = 'false'
        if ($script:Locked) { $lck = 'true' }
        $hb = '{"pid":' + $PID +
              ',"startedAt":"' + $script:StartedAt.ToString('yyyy-MM-dd HH:mm:ss') + '"' +
              ',"uptimeSec":' + $up +
              ',"expanded":' + $exp +
              ',"locked":' + $lck +
              ',"at":"' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '"}'
        Write-CtlFile 'status.json' $hb
    } catch {}
}

function Invoke-CtlCommand([string]$Verb, [string]$Arg) {
    switch ($Verb) {
        'stop'     { Write-CtlResult 'stop' 'ok'; Invoke-BeaconStop }
        'expand'   { if (-not $w.IsVisible) { $w.Show() }; Expand-Orb; Write-CtlResult 'expand' 'ok' }
        'collapse' { Collapse-Orb; Write-CtlResult 'collapse' 'ok' }
        'lock'     { $script:Locked = $true; Update-LockUI; Write-CtlResult 'lock' 'ok' }
        'unlock'   { $script:Locked = $false; Update-LockUI; Write-CtlResult 'unlock' 'ok' }
        'refresh'  { Reload-Notes; if ($script:IsOpen) { Render-Classes; Render-Notes }; Write-CtlResult 'refresh' 'ok' }
        'dump'     { Dump-AgendaUi; Write-CtlResult 'dump' 'ok' }
        'ping'     { Write-CtlResult 'ping' 'ok' }
        default    { Write-CtlResult $Verb 'unknown' }
    }
}

$ctlTimer = New-Object System.Windows.Threading.DispatcherTimer
$ctlTimer.Interval = [TimeSpan]::FromSeconds(2)
$ctlTimer.Add_Tick({
    Write-BeaconHeartbeat
    $cmdFile = Join-Path $CtlDir 'cmd.txt'
    if (-not (Test-Path $cmdFile)) { return }
    $lines = @()
    try { $lines = @(Get-Content -LiteralPath $cmdFile -ErrorAction Stop) } catch { return }
    Remove-Item -LiteralPath $cmdFile -Force -ErrorAction SilentlyContinue
    foreach ($line in $lines) {
        $text = ([string]$line).Trim()
        if (-not $text) { continue }
        $parts = @($text -split '\s+', 2)
        $verb  = ([string]$parts[0]).ToLower()
        $arg   = ''
        if ($parts.Count -gt 1) { $arg = [string]$parts[1] }
        try { Invoke-CtlCommand $verb $arg } catch {
            try { Write-CtlResult $verb ('error: ' + $_.Exception.Message) } catch {}
        }
    }
})
$ctlTimer.Start()

Write-BeaconHeartbeat

# schedule refresh every 5min
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMinutes(5)
$timer.Add_Tick({
    Reload-Notes
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

# A crash used to leave nothing behind. Log it (do not swallow: behaviour stays the same).
$app.Add_DispatcherUnhandledException({
    param($s, $e)
    try {
        $line = "[{0}] {1}`r`n{2}`r`n`r`n" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'),
                $e.Exception.GetType().FullName, $e.Exception.ToString()
        Add-Content -Path (Join-Path $Base 'error.log') -Value $line -Encoding UTF8
    } catch {}
})

$app.Run($w) | Out-Null
