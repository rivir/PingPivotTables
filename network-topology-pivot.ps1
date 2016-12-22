## NETWORK TOPOLOGY INTERCONNECTION STATS  v.1
##  Thomas Marantz
##
##  Test-Connection + Pivot Tables + HTML Export (jquery.hottie.js HeatMap)
##		Powershell 3
##		PowerShell Pivots https://gist.github.com/andyoakley/1651859
##		JQuery https://cdnjs.cloudflare.com/ajax/libs/jquery/1.8.3/jquery.js
##		Hottie https://github.com/DLarsen/jquery-hottie
##		Pure CSS http://purecss.io/ https://unpkg.com/purecss@0.6.1/build/pure-min.css
##  December 2016
##
##
###########################################

#Computers that will ping out to gather ping data. Must have Powershell v3 with Remote
#$psversiontable
#get-service winrm
#Enable-PSRemoting –force
[System.Array]$arrSource = @('HAWAIISERVER','SANFRANCISCO','NEWYORK')

#Computers that will be pinged
[System.Array]$arrComputer = @('HAWAIISERVER','SANFRANCISCO','NEWYORK')

##Ping all systems using Test-Connection cmdlet
$all_pings = @()

Foreach ($x in $arrSource){
        workflow TestAllTheConnections{
            param(
                [System.String[]]$Source,
                [System.String[]]$Computers
            )
                Foreach -Parallel ($c in $Computers){               
                        Test-Connection -Source $Source -ComputerName $c -Count 7 -ErrorAction SilentlyContinue | Select -Skip 1            
              }
            
            }
        $all_pings += TestAllTheConnections -Source $x -Computers $arrComputer 
    }

##################
# Create variable to store the output
$rows = @()
### FIND OUTAGE SITES

#array with all sites that were pinged
$all_contacted = $all_pings | Select -unique Address

#Pivot table data
$data = $all_pings | Select __SERVER,Address,ResponseTime

##Find all sites that did not respond to ping
$arrComputer |?{($_) -notin  ($all_contacted | Select -exp Address)} | ForEach-Object { 

    Foreach ($c in $arrSource){               
                        $data += [PSCustomObject]@{__SERVER=$c;Address=$_;ResponseTime="NA"}                
              }
    
    }


##PIVOT TABLE 1
$data = $data | Sort-Object Address
# Fields of interest
$keep = "Address" # Bits along the top
$rotate = "__SERVER" # Those along the side
$value = "ResponseTime" # What to total

# Find the unique "Rotate" [top row of the pivot] values and sort ascending
$pivots = $data | select -unique $rotate | foreach { $_.$rotate} | Sort-Object

# Step through the original data...
#  for each of the "Keep" [left hand side] find the Sum of the "Value" for each "Rotate"

$data | 
    group $keep | 
    foreach { 
        $group = $_.Group
        # Create the data row and name it as per the "Keep"
        $row = new-object psobject
        $row | add-member NoteProperty $keep $_.Name 
        # Cycle through the unique "Rotate" values and get the sum
        foreach ($pivot in $pivots)
        { 
            $row | add-member NoteProperty $pivot ($group | where { $_.$rotate -eq $pivot } | measure -Minimum $value).Minimum
        }
        # Add the total to the row
        #$row | add-member NoteProperty Total ($group | measure -sum $value).Sum
        # Add the row to the collection 
        $rows += $row
    } 

##PIVOT TABLE 2
##MAXIMUM
$rows2 = @()
 $data | 
    group $keep | 
    foreach { 
        $group2 = $_.Group
        # Create the data row and name it as per the "Keep"
        $row2 = new-object psobject
        $row2 | add-member NoteProperty $keep $_.Name 
        # Cycle through the unique "Rotate" values and get the sum
        foreach ($pivot in $pivots)
        { 
            $row2 | add-member NoteProperty $pivot ($group2 | where { $_.$rotate -eq $pivot } | measure -Maximum $value).Maximum
        }
        # Add the total to the row
        #$row | add-member NoteProperty Total ($group | measure -sum $value).Sum
        # Add the row to the collection 
        $rows2 += $row2
    } 

##PIVOT TABLE 3
##AVERAGE
$rows3 = @()
 $data | 
    group $keep | 
    foreach { 
        $group3 = $_.Group
        # Create the data row and name it as per the "Keep"
        $row3 = new-object psobject
        $row3 | add-member NoteProperty $keep $_.Name 
        # Cycle through the unique "Rotate" values and get the sum
        foreach ($pivot in $pivots)
        { 
            if (($group3 | where {$_.$rotate -eq $pivot} | select -exp ResponseTime) -eq "NA") {
               $row3 | add-member NoteProperty $pivot ($group3 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime)
            } else {
               $avgping = ($group3 | where { $_.$rotate -eq $pivot } | measure -Average $value).Average
               $row3 | add-member NoteProperty $pivot ("{0:N0}" -f $avgping)
            }

        }
        # Add the total to the row
        #$row | add-member NoteProperty Total ($group | measure -sum $value).Sum
        # Add the row to the collection 
        $rows3 += $row3
    } 


##PIVOT TABLE 4 JITTER
$rows4 = @()
 $data | 
    group $keep | 
    foreach { 
        $group4 = $_.Group
        # Create the data row and name it as per the "Keep"
        $row4 = new-object psobject
        $row4 | add-member NoteProperty $keep $_.Name 
        # Cycle through the unique "Rotate" values and get the sum
        foreach ($pivot in $pivots)
        { 
            if (($group4 | where {$_.$rotate -eq $pivot} | select -exp ResponseTime) -eq "NA") {
               $row4 | add-member NoteProperty $pivot ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime)
            } else {
               $jitter = 0
               #three jitter calcs
               $jitter = [math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 1))
               $jitter += [math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 1) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 2))
               $jitter +=[math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 2) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 3))
               $jitter += [math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 3) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 4))
               $jitter +=[math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 4) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 5))
               $jitter += [math]::abs(($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 5) - ($group4 | where { $_.$rotate -eq $pivot } | select -exp ResponseTime -First 1 -skip 6))
               $jitter = $jitter/6
               $row4 | add-member NoteProperty $pivot ("{0:N0}" -f $jitter)
            }

        }
        # Add the total to the row
        #$row | add-member NoteProperty Total ($group | measure -sum $value).Sum
        # Add the row to the collection 
        $rows4 += $row4
    } 

#Html Header, Jquery 1.8.3, CSS
$Header = @"
 <meta http-equiv="refresh" content="600">
 <meta name="viewport" content="width=device-width, initial-scale=1">
  <style type="text/css">
    .hover { background-color:rgba(0, 0, 0, 0.5); }
    ul { margin-bottom: 10px !important; }
    table { margin: 8px !important; }
    td { padding: 8px !important; text-align: center; }
    * {font-family: sans-serif;font-size: 14px; background-color: #F0F0F0;}
    .child_div{
       float: left;
    }
    .parent_div{
        clear:both;
    }
  </style>

<script type="text/javascript" src="jquery-1.8.3.js"></script>
<link rel="stylesheet" href="pure-min.css">
<script type="text/javascript" src="jquery.hottie.js"></script>

<script type='text/javascript'>//<![CDATA[
`$(window).load(function(){

        `$("td").hottie({
        colorArray : [ 
            "#63BE7B", // highest value
            "#FCCFCF",
             "#F8696B" // lowest value
            // add as many colors as you like...
        ],
        nullColor : "#FA4848"
    });

    `$("div#jitter_table td").hottie({
        colorArray : [ 
            "#99CCFF", // highest value
            "#666666",
             "#FF99CC" // lowest value
            // add as many colors as you like...
        ],
        nullColor : "#FA4848"
    });



var txt = 'Address';
var column = `$('table tr th').filter(function() {
    return `$(this).text() === txt;
}).index();

if(column > -1) {
    `$('table tr').each(function() {
        `$(this).find('td').eq(column).css('background-color', '#FFFFFF');
    });
}

`$("table").delegate('td','mouseover mouseleave', function(e) {
        if (e.type == 'mouseover') {
          `$(this).parent().addClass("hover");
          `$("colgroup").eq(`$(this).index()).addClass("hover");
        } else {
          `$(this).parent().removeClass("hover");
          `$("colgroup").eq(`$(this).index()).removeClass("hover");
        }
    });

});//]]> 

</script>

<script>

`$(document).ready(function(){

    `$('#selector').on('change', function() {

        `$('#min_table').hide();
        `$('#max_table').hide();
        `$('#avg_table').hide();
      if ( this.value == '2')
      {
        `$('#min_table').show();
      }
      else if ( this.value == '3')
      {
        `$('#max_table').show();
      }
      else
      {
        `$('#avg_table').show();
      }
    });
});

</script>

<title>KJC Network Pings Table</title>


"@
 



# HTML Output
$table1 = $rows | ConvertTo-HTML -Fragment -PreContent '<div class="pure-u-1-2" id="min_table"><h2>Minimum Ping Times(ms)</h2>’ -PostContent '</div></div>' | Out-String

$table2 = $rows2 | ConvertTo-Html -Fragment -PreContent '<div class="pure-g"><div class ="pure-u-1-2" id="max_table"><h2>Maximum Ping Times(ms)</h2>’ -PostContent '</div>' | Out-String

$table3 = $rows3 | ConvertTo-Html -Fragment -PreContent '<div class ="pure-u-1-2" id="avg_table"><h2>Average Ping Times(ms)</h2>’ -PostContent '</div></div>' | Out-String

$table4 = $rows4 | ConvertTo-Html -Fragment -PreContent '<div class="pure-g"><div class ="pure-u-1-2" id="jitter_table"><h2>Average Jitter Times(ms)</h2>’ -PostContent '</div>' | Out-String

$closing = "</div>"


$precc = @"
<div class="pure-g">
    <div class="pure-u-1-3">$((Get-Date).ToString()) run via $env:computername</div>
    <div class="pure-u-1-3">Table Selector:<select id="selector"><option value=1 selected="selected">Average</option><option value=2 >Minimum</option><option value=3 >Maximum</option></select></div></div>
<div class="pure-menu pure-menu-horizontal pure-menu-scrollable">
    <ul class="pure-menu-list">
        <li class="pure-menu-item"><a href="http://server/pings.html" class="pure-menu-heading pure-menu-link">HOME</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.02.html" class="pure-menu-link">2</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.03.html" class="pure-menu-link">3</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.04.html" class="pure-menu-link">4</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.05.html" class="pure-menu-link">5</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.06.html" class="pure-menu-link">6</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.07.html" class="pure-menu-link">7</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.08.html" class="pure-menu-link">8</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.09.html" class="pure-menu-link">9</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.10.html" class="pure-menu-link">10</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.11.html" class="pure-menu-link">11</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.12.html" class="pure-menu-link">12</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.13.html" class="pure-menu-link">13</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.14.html" class="pure-menu-link">14</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.15.html" class="pure-menu-link">15</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.16.html" class="pure-menu-link">16</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.17.html" class="pure-menu-link">17</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.18.html" class="pure-menu-link">18</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.19.html" class="pure-menu-link">19</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.20.html" class="pure-menu-link">20</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.21.html" class="pure-menu-link">21</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.22.html" class="pure-menu-link">22</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.23.html" class="pure-menu-link">23</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.24.html" class="pure-menu-link">24</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.25.html" class="pure-menu-link">25</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.26.html" class="pure-menu-link">26</a></li>
        </ul>
        <ul class="pure-menu-list">
        <li class="pure-menu-item"><a href="http://server/ping/pings.27.html" class="pure-menu-link">27</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.28.html" class="pure-menu-link">28</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.29.html" class="pure-menu-link">29</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.30.html" class="pure-menu-link">30</a></li>        
        <li class="pure-menu-item"><a href="http://server/ping/pings.31.html" class="pure-menu-link">31</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.32.html" class="pure-menu-link">32</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.33.html" class="pure-menu-link">33</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.34.html" class="pure-menu-link">34</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.35.html" class="pure-menu-link">35</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.36.html" class="pure-menu-link">36</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.37.html" class="pure-menu-link">37</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.38.html" class="pure-menu-link">38</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.39.html" class="pure-menu-link">39</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.40.html" class="pure-menu-link">40</a></li>        
        <li class="pure-menu-item"><a href="http://server/ping/pings.41.html" class="pure-menu-link">41</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.42.html" class="pure-menu-link">42</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.43.html" class="pure-menu-link">43</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.44.html" class="pure-menu-link">44</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.45.html" class="pure-menu-link">45</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.46.html" class="pure-menu-link">46</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.47.html" class="pure-menu-link">47</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.48.html" class="pure-menu-link">48</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.49.html" class="pure-menu-link">49</a></li>
        <li class="pure-menu-item"><a href="http://server/ping/pings.50.html" class="pure-menu-link">50</a></li>
    </ul>
</div>
"@


#RENAME FILES
#Delete pings.50.html, increment 02..50

$nr=50
$folder = "\\server\wwwroot\ping\"
Get-ChildItem $folder -Filter "pings.50.html" | Remove-Item
Get-ChildItem $folder -Filter "pings.*.html" | Sort-Object name -descending | % {Rename-Item $folder$_ -NewName ('pings.{0}.html' -f $nr--)} 
$nr=9
Get-ChildItem $folder | Where {$_.name -match "pings\.[0-9]\.html"} | Sort-Object name -descending | % {Rename-Item $folder$_ -NewName ('pings.0{0}.html' -f $nr--)} 

#COPY FILE
Copy-Item \\server\wwwroot\pings.html \\server\wwwroot\ping\pings.02.html 

#OUTPUT NEW TABLE
ConvertTo-HTML -Head $Header -PostContent $table3,$table1,$table4,$table2,$closing `
-PreContent  $precc `
 | Out-File \\server\wwwroot\pings.html