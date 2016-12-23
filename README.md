# PingPivotTables
Use Powershell Test-Connection to have Servers pinging servers and display ping times in a table.

- PS Test-Connection + Pivot Tables + HTML Export (jquery.hottie.js HeatMap)

<img src="screenshot.png" align="center" />

## Getting Started

Download repository. Update paths to your local web server. Use Task Scheduler to run script as often as desired.

### Prerequisites

PowerShell v3
PowerShell Remoting turned on pinging systems

```
$psversiontable
get-service winrm
Enable-PSRemoting –force
```

### Configuring

Update paths on network-topology-pivot.ps1
```
<script type="text/javascript" src="jquery-1.8.3.js"></script>
<link rel="stylesheet" href="pure-min.css">
<script type="text/javascript" src="jquery.hottie.js"></script>
```

find/replace //server/ping/ to your web server location
```
<li class="pure-menu-item"><a href="http://server/ping/pings.02.html" class="pure-menu-link">2</a></li>
```

```
$folder = "\\server\wwwroot\ping\"
Copy-Item \\server\wwwroot\pings.html \\server\wwwroot\ping\pings.02.html 
Out-File \\server\wwwroot\pings.html
```

## Built With

- Microsoft PowerShell 3
- [JQuery] (https://cdnjs.cloudflare.com/ajax/libs/jquery/1.8.3/jquery.js)
- [Hottie] (https://github.com/DLarsen/jquery-hottie)
- [Pure CSS] (http://purecss.io/) https://unpkg.com/purecss@0.6.1/build/pure-min.css

## Authors

* **Tom Marantz** - *Initial work*

## Acknowledgments

- [PowerShell Pivots] (https://gist.github.com/andyoakley/1651859)
