function Get-CBLyncHealthNew {

  <#
    .SYNOPSIS
    Check the health of a Lync/Skype server farm.

    .DESCRIPTION


    .EXAMPLE
    run the script without parameters and it will use the information already found within the Lync topology

    .NOTES
    Place additional notes here.

    .LINK
    https://www.linkedin.com/in/cburns/

    .INPUTS
    No Inputs needed

    .OUTPUTS
    Will ouput a HTML file to be used in Reports (PShtml.html)
  #>
  [cmdletbinding()]
   Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $Server,
    # CSS location - please enter the location of the css code to format your table
    [Parameter(Position=1)]
    [String]
    $CSSFile,
    # User Logo, this will alter the logo used in the top left of the report
    [Parameter(Position=2)]
    [String]
    $UserLogo = "https://upload.wikimedia.org/wikipedia/commons/8/86/Microsoft_Skype_for_Business_logo.png",
    # Federation Partner Test
    [Parameter(Position=3)]
    [String[]]
    $FederationPartners = @('dell.com','gcicom.net','circleit.co.uk')

  )

#region Variables

$CompanyLogoBase64 = (Get-CBlyncBase64 -Url $UserLogo)
$colorscheme = '#222222'
$brand2color = '#425563'
$overrideAutoStrip = $false
$charstokeepforserverid =2


#region Css

$Header = @"

<meta name='viewport' content='width=device-width, initial-scale=1.0'><style>
 /* This section is for the main interface - navigation etc */

.nav {
    background-color: #111111;
    color: white;


    width: 200px;
    height: 100hv;
}

div.logo {
    background-color: white;
    color: white;
    position: absolute;
    left: 0px;
    top: 0px;


    padding: 20px;
    padding-left: 30px;
}

div.content {

    background-color: #eeeeee;

    height: 100hv;
    padding-top: 20px;

    flex-grow: 1;
    -ms-flex: 1;
}

.mainframe {
    background-color: #d9d9d9;
    width: 100%;
    display: flex;
    display: -ms-flexbox;
    border-top-style: solid;
    border-top-color: #666666;
    border-top-width: 1px;
    left: 0px;
    top: 100px;
    height: 100%;
    position: absolute;
}

div.navbuts {
    padding-top: 20px;

    width: 100%;
    position: relative;
}

div.navbut {
    text-color: white;
    background-color: #111111;
    position: relative;
    border-left-style: solid;
    border-left-color: #111111;
    border-left-width: 2px;
    padding: 10px;
    padding-right: 20px;
    text-align: right;
    font-family: "arial", arial, sans-serif;
    margin-top: 8px;
    border-right-style: solid;
    border-right-color: #ffed00;
    border-right-width: 0px;
    display: flex;
    display: -ms-flexbox;
}

div.navbut:hover {
    border-left-color: #ffed00;
    cursor: pointer;
}

div.navbut:hover+.ContentDiv {
    display: block;
}

.ContentDiv {
    background-color: #d9d9d9;
    padding: 0px;
    display: block;
    margin-left: 40px;
    margin-top: 15px;
    margin-right: 40px;
    font-family: "arial", arial, sans-serif;
    box-shadow: 0 4px 8px 0 rgba(0, 0, 0, 0.2), 0 6px 20px 0 rgba(0, 0, 0, 0.19);
    border-top-right-radius: 3px;
    border-top-left-radius: 3px;
    border-bottom-right-radius: 3px;
    border-bottom-left-radius: 3px;
}

.ContentDiv h1 {

    color: #000000;
    padding: 10px;
    padding-left: 20px;
    font-size: 18pt;
    border-bottom-style: solid;
    border-bottom-color: #ffed00;
    border-bottom-width: 3px;
}

.ContentInnerDiv {
    padding: 30px;
    padding-top: 0px;
}

#Overview {
    display: block;
}

.IssueCount {

    padding: 20px;
    width: 250px;
    height: 61px;
    position: absolute;
    right: 20px;
    top: 0px;

    border-left-style: solid;
    border-left-color: #d9d9d9;
    border-left-width: 1px;
}

.IssueCountNumber {
    background-color: #ffed00;
    font-size: 25px;
    text-align: center;
    padding: 10px;
    width: 40px;
    height: 30px;
    margin: 5px;
    font-family: "arial", arial, sans-serif;
    float: left;
}

.issuesTitle {
    margin: 20px;
    font-size: 20;
    align: center;
    margin-left: 85px;
    font-family: "arial", arial, sans-serif;
}

.tooltip {
  position: relative;
  display: inline-block;
  border-bottom: 1px dotted black;
}

.tooltip .tooltiptext {
  visibility: hidden;
  width: 250px;
  background-color: #555;
  color: #fff;
  text-align: center;
  border-radius: 6px;
  padding: 5px 0;
  position: absolute;
  z-index: 1;
  top: 130%;
  left: 1%;
  margin-left: -60px;
  opacity: 0;
  transition: opacity 0.3s;
}
.tooltip .tooltiptext table{
      width: 100%;
      border-width: 1px;
    border-style: solid;
    border-color: black;
    border-collapse: collapse;
    table-layout: fixed;
}

.tooltip .tooltiptext TH {
    border-width: 1px;
    padding: 3px;
    border-style: solid;
    border-color: black;
    font-size: 10pt;
    color: white;
    background-color: #425563;
}

.tooltip .tooltiptext TD {
    border-width: 0px;
    border-style: solid;
    border-color: black;
    font-size: 8pt;
    vertical-align: top;
    padding: 3px
}

  .tooltip .tooltiptext TR:nth-child(odd) {
    background-color: gray
}

.tooltip .tooltiptext::after {
  content: "";
  position: absolute;
  bottom: 100%;
  left: 50%;
  margin-left: -5px;
  border-width: 5px;
  border-style: solid;
  border-color: transparent transparent #555 transparent;
}

.tooltip:hover .tooltiptext {
  visibility: visible;
  opacity: 1;
}


img.logo {
    
}

.reportdate {
    font-size: 10px;
    font-family: "arial", arial, sans-serif;
    text-align: right;
    top: 105px;
    right: 15px;
    position: absolute;
}

.pooltable {
    padding: 15px;
}


@media screen and (max-width: 1150px) {
    .nav {
        width: 60px;
        height: 100hv;
    }
    .poolgraphic {
        position: absolute;
        clip: rect(0px, 300px, 300px, 0px);
        height: 250px;
    }
    div.mainframe {
        top: 80px;
    }
    .ContentDiv {
        display: block;
        margin-left: 20px;
    }
    .IssueCount {
        display: none;
    }
    div.content {
        margin-left: 50px;

        width: 97%;
    }
    img.logo {
        height: 45px;
    }
    .menuword {
        display: none;
    }
    div.navbuts {
        padding-top: 0px;
        width: 100%;
        position: relative;
    }
    .reportdate {
        top: 85px;
    }
}
/* This section is for the main interface - navigation etc */
/* This section controls the database graphics to show which side is primary vs Mirror */

.SQLBalance {
    position: relative;
}

tr {
    text-align: center;
    border: none;
    font-family: "arial", arial, sans-serif;
}

.Primaryside {
    background-color: green;
    width: 300px;
    border: 0px;
}

.Mirrorside {
    background-color: black;
    width: 200px;
    border: 0px;
}


.normal-right {

    width: 0px;
    height: 0;

    border-top: 20px solid black;
    border-bottom: 20px solid black;

    border-left: 20px solid green;
}

.logo {
    background-color: #ffffff;
    
    height: 70px;
}

.normal-left {

    width: 0px;
    height: 0;

    border-top: 20px solid black;
    border-bottom: 20px solid black;

    border-right: 20px solid green;
}

.failed-right {

    width: 0px;
    height: 0;

    border-top: 20px solid red;
    border-bottom: 20px solid red;

    border-left: 20px solid green;
}

.failed-left {

    width: 0px;
    height: 0;

    border-top: 20px solid red;
    border-bottom: 20px solid red;

    border-right: 20px solid green;
}

.stattable {
    margin: 20px;
    height: 50px;
    color: white;
}

.menuword {
    flex-grow: 1;
    -ms-flex: 1;
}

.menuicon {
    width: 30px;
}

.menuimg {
    Width: 20px;
    height: 20
}





TABLE {
    border-width: 1px;
    border-style: solid;
    border-color: black;
    border-collapse: collapse;
    table-layout: fixed;
}

TH {
    border-width: 1px;
    padding: 3px;
    border-style: solid;
    border-color: black;
    font-size: 10pt;
    color: white;
    background-color: $brand2color;
}

TD {
    border-width: 0px;
    border-style: solid;
    border-color: black;
    font-size: 8pt;
    vertical-align: top;
    padding: 10px
}

TR:nth-child(odd) {
    background-color: lightgray
}

.pools td:nth-child(1) {
    color: $colorscheme;
}

.pools td:nth-child(2) {
    color: $brand2color;
}

.computer_pass {
    border-width: 2px;
    border-style: solid;
    border-color: black;
    vertical-align: top;
    padding: 10px;
    color: white;
    background-color: green;
}

.computer_fail {
    border-width: 2px;
    border-style: solid;
    border-color: black;
    vertical-align: top;
    padding: 10px;
    color: white;
    background-color: red;
}

.replica_pass {
    border-width: 2px;
    border-style: solid;
    border-color: black;
    vertical-align: top;
    padding: 10px;
    color: white;
    background-color: green;
}

.replica_fail {
    border-width: 2px;
    border-style: solid;
    border-color: black;
    vertical-align: top;
    padding: 10px;
    color: white;
    background-color: red;
}


.FEandMed_container {
    display: flex;
}

.EdgeandSQL_Container {
    display: flex;
}

.Replication_Container {
    display: flex;
}

.servers {
    min-height: 300px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    min-width: 400px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.servers h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.servers h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.serverdivcontent {
    padding: 10px;
}

.pooldist {
    min-height: 450px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    min-width: 500px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.pooldist h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.pooldist h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.poolinfo {
    min-height: 300px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.poolinfo h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.poolinfo h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.usercountinfo {
    min-height: 450px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    width: 200px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.usercountinfo h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.usercountinfo h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.totalusers#Ttlusers {
    background:   /* on "bottom" */
                    
          -webkit-linear-gradient(
              315deg,
              green, darkgreen);
                    background:   /* on "bottom" */
                    
          -moz-linear-gradient(
              315deg,
              green, darkgreen);
                    background:   /* on "bottom" */
                    
          -o-linear-gradient(
              315deg,
              green, darkgreen);
                    background:   /* on "bottom" */
                    
          linear-gradient(
              315deg,
              green, darkgreen);


;
}
.totalusers#Ttlusers h3{
    background: darkgreen;
}


.totalusers#EVusers {
    background:   /* on "bottom" */
                    
          -webkit-linear-gradient(
              315deg,
              orchid, darkorchid);
                    background:   /* on "bottom" */
                    
          -moz-linear-gradient(
              315deg,
              orchid, darkorchid);
              background:    /* on "bottom" */
                    
          -o-linear-gradient (
              315deg,
              orchid, darkorchid);
         background:   /* on "bottom" */
                    
          linear-gradient(
              315deg,
              orchid, darkorchid);


;

}
.totalusers#EVusers h3{
    background: darkorchid;
}


.totalusers#ExUMusers {
    background:   /* on "bottom" */
                    
          -webkit-linear-gradient(
              315deg,
              salmon, darksalmon);
                    background:   /* on "bottom" */
                    
          -moz-linear-gradient(
              315deg,
              salmon, darksalmon);
                    background:   /* on "bottom" */
                    
          -o-linear-gradient(
              315deg,
              salmon, darksalmon);
                    background:   /* on "bottom" */
                    
          linear-gradient(
              315deg,
              salmon, darksalmon);


;
}
.totalusers#ExUMusers h3{
    background: darksalmon;
}

.totalusers{
                    color: white;
                    padding: 0px;
                    margin: 20px;
                    font-size: 18pt;
                    height: 100px;
                    width: 80%;
                }
.totalusers h3 {
   text-align: left;
   
   color: white;
    padding: 3px;
    margin: 0;
    font-size: 10pt;
}

.totalusersspan {
   text-align: right;
   padding: 5px;
   font-size: 34pt;
   width: 95%;
   Padding-top: 20px;
}

.repservices {
    min-height: 300px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    min-width: 400px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.repservices h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.repservices h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.sqlservices {
    min-height: 300px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    width: 280px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.sqlservices h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.sqlservices h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}

.addressservices {
    min-height: 300px;
    border-width: 2px;
    border-style: solid;
    border-color: $brand2color;
    vertical-align: top;
    margin: 10px;
    min-width: 400px;
    display: inline-block;
    position: relative;
    border-top-right-radius: 7px;
    border-top-left-radius: 7px;
    border-bottom-right-radius: 2px;
    border-bottom-left-radius: 2px;
    background-color: #F0F0F0
}

.addressservices h2 {
    background: $brand2color;
    color: white;
    padding: 10px;
    margin: 0;
    font-size: 12pt;
}

.addressservices h3 {
    padding: 0px;
    margin: 0px;
    font-size: 10pt;
    margin-top: 5px
}


.federationservices {
  min-height: 300px;
  border-width: 2px;
  border-style: solid;
  border-color: $brand2color;
  vertical-align: top;
  margin: 10px;
  min-width: 200px;
  display: inline-block;
  position: relative;
  border-top-right-radius: 7px;
  border-top-left-radius: 7px;
  border-bottom-right-radius: 2px;
  border-bottom-left-radius: 2px;
  background-color: #F0F0F0
}

.federationservices h2 {
  background: $brand2color;
  color: white;
  padding: 10px;
  margin: 0;
  font-size: 12pt;
}

.federationservices h3 {
  padding: 0px;
  margin: 0px;
  font-size: 10pt;
  margin-top: 5px
}


.servergraphic {
    position: absolute;
    bottom: -10px;
    right: -20px;
    height: 50px
}

.subtitle {
    font-size: 8pt;
    font-style: oblique;
}



div#poolinfo:target {
    display: block;
}
</style>
"@

$JavaScript =@"

<script>
function InfoPanes(toShow,button){
    var x = document.getElementsByClassName('ContentDiv');
    var view = 'none';
    if(toShow == 'ALL'){
        view = 'block'
    }
    for(i=0; i < x.length; i++)
    { 
        x[i].style.display=view;
    };
    var y = document.getElementsByClassName('navbut');
    for(i=0; i < y.length; i++)
    { 
        y[i].style.borderRightWidth='0px';
        y[i].style.color='white';
    };
    var currentButton = document.getElementById(button);
    currentButton.style.borderRightWidth='5px';
    currentButton.style.color='#ffed00'
    if(toShow != 'ALL'){
        document.getElementById(toShow).style.display='block';
    }
    return false;
};
document.addEventListener('DOMContentLoaded', function HideAllButOVW(){
var x = document.getElementsByClassName('ContentDiv');
for(i=1; i < x.length; i++)
    { 
        x[i].style.display='none';
    };

},true);
</script>

"@
#endregion


#region Graphics
$EdgeGraphicBase64 = 'data:image/png;base64,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'
$FrondEndGraphicBase64 = 'data:image/png;base64,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'
$MediationGraphicBase64 = 'data:image/png;base64,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'
$DBGraphicBase64 = 'data:image/png;base64,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'
$ReplicationGraphicBase64 = 'data:image/png;base64,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'
$OverviewGraphic = 'data:image/png;base64,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'
$CapacityGraphic = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAAAZy0lEQVR42u2de5TdVXXHP7kZx+l0HNOskMYY0zTSFCGF1BUgZaE9TQEBEUUqL3mJvETloY2IlMVmIVKMGhB5KKC8Hz4AEREpwq8RY4yYpjGmIcasdFYMmMasMWs6jtPLpX+c85M777n3/u69v/v7fT9r3TWBZObe2b+zv2efvfc5ZwoTYGbTgPlAd3i1AV2j/NNOoJ3qKAF7xvi7ItBX9u/6hv3/gfDqAwbNbA9C1AEzmw7MAGaG1wygY5J+MAj0j/Pj94TxPZnvKR/j8fgf/u/7Yh8xs+JYbzplDIc/HTgcWBJ+yVYiFom+YIxeYFd47QReBF4CdpS99piZRnj+HDr+4zRgDjAbmAu8KTj4nPB1dvja3qK/ai/QA9wB3GpmgyMEIBjjJOC2MWb4rFIKwrAlvH4JbAY2AVtjRZVAtLSDd4codu/w9S+Do88F5uVsvD8FHGdm/cMF4Ajge0BBQ2fI8mMbsDEIwn8Ca4NQFCUKqXP26cD+wCJgP2BBeM3UuB7CKuDvzaw4JRivALwQFFJMzE4gAh4CnjCzAZmkaY4/C/gwcCywLz5HJSbmCjP7dCwAS4AfyyZVsQY4zcw2yxQNn/GXAl8LobyojAHgDXFYdJTsUTUHAd8xs31kioY6/ynAI3L+qukAjokFYJHsURMLgPvMbKZM0RCWArfgk3uieubHArBQtqiZRfjyqajv7N8GXCHnT4ZCCKfmyBS12xI4QGaoO9PwpTuR0KCdRes2OKTRnqK+lGSC5AVAJEOfTNAQGw/KDMkJwHSZITF+LRPUPQcwyNj7RkQVAtAlMyTGTpmgIWyXCZITAK1bk2El8LDM0BBW4Dd0CUUAqXH+95mZIoDGLANWAidKBBQByPklAhIBCYCcXyIgqhUAdVTJ+SUCORYAIeeXCGgJICZJJOeXCGgJkF/nP1HOLxHQEkDOLyQCEgA5v5AISACySknOLxGQAOTX+VfK+VteBE6WCEgAqmG1nD8TIhBJBCQAlbJGzp85ETgNfzuUGCYA2ls9lPXB+bXlNFsi8DRwBrBb1hgqADpi6VU2AMeb2TaZIpMi8AT+EpFeWUNLgOFsxnf4bZEpMi0CDwKXoOPbJABl9AAnm9kmmSIXInAncBnjX9ctAcgJg/h70tbKFLniZuAxCYBCoXbg8HDhhMgPc/HXuuVeAIoaC5wAnKXrvnOzBGgHlgPzJQAijgKuAZbIFJl3/gLwEfx14kgARMwM4MtmNlumyDQOf7egbsNSDmAE+wM3mVmnTJHJ2X8e/lbhabKGcgBjcQxwpZKCmXP+ruD8C2QNLQHGow24EDhTScHMOH8bcBVwhKwxUgD6ZYYRdADXAUtlipZ3foBTgQs04WkJUAnTgdvMbB+ZoqVZgi/5dcgUWgJUynzgDjObKVO05Ow/F/gyvsIjtASoikNQZaAVnb8LuAlf2RHjCMCgzDAh7wGuVmWgZZw/TvodKWtMLADqA5iYNnwS6dzQSSbS6/wF4KzwvCTYigASowO4Fjha5cFU8/bwnJT0m6QA6EiwydONbyZZLFOkcvZfgE/6TZc1Ji8AKgNWxhx8ZWC+TJEq558enF+dfpUIgJkpB1A5++N7BFReSofzd+Br/W+XNSqPAAAGZIqKccCKMPhE85y/AFyM7/ZTgrZKAVAvQHW2OwltHGqm8wO8F7gcbe+tSQC0DKiOtjD7nK3yYFNYDNwAdMkUWgI0i7g8eIzKgw2d/ecCtwE6wCUBAVApsDam4dtOD5IpGuL80/Dl2EWyhpYAaSEuD+4tU9TV+dtRm2/iAqCrkpJhIf5cQe0erI/zF4DzgXNRxj8J+iUAyePw5UElppJ1foCjw+yv0msy7I4FQDemJhtVnQBcFcJVkQyL8HkWHeiZHDtiAXhRtkiUePfgR1QeTGT2n4Nv850rayTK9nhwviRbJE4HcCVwgsqDNTl/F77WrwpLHQVgu2xRF7qBFcChMkVVzt+Gv8RDt/gkT6+Z9cYCsE32qBuz8JUB7VKrzPkLwJn4a7zUap082+DVUsoOtC24nuwbRGCWTDFp4oM9dBZjfdhaLgCDQQSaQT9QysmAvkHlwUnN/gvwGf+sb7fub+LE+2oEEJJUPU1SoROBT5P9duQCfueadg+O7/zT8Em/fTP8a5aAdcD7gM/QnN24vyqPAJqRB9gAvNvMHsc3dxwHbMz4+G4La9qzVB4c1fnbgavJ9hVeg8D9wDvM7Al8pejjTZgAtw4XgF822PmPM7MN4cGXzOwZ4Cjg0YznIzqAa9C1Y8OdPz7N92yy2+bbB1wGfMDMdsbRt5ndCnyIxnbkjhCALQ164/XB+beMMgh6gJPDLJDlQ0pmADeqMjAiR3I12W3z3R7G9vVmVhxl7N8PfBDY1aAoZNtwAdjcgDdeE8L+LePMBAMhJ/D+JuUlGsU+QQRy39pqZvPCuj+rSb/nQ3T7uJmVxrHDw8AZwM46f56eIAIjIoB6ZeNLwHPA8Wa2bRIDohSWAkeF78tqleAwYFme8wHhyrXlZPMKryLwTeBdZrZhMh2hIS9wIvXtzt0Sf5bygddbp/CjBKwMzr+9goGBmW3EJwfvJJsXmBTwewbenlPnj7f3ZrHTbwD4HHCGmb1UoV2isFyolwhsKh+Ao/5FgqwEToyTHlUMkl0hQXIp2SwVTgMuzenOwYX4DHjWfvde4BLgCjPrr3LcR/gyYT2Wwb8YIQAhJEhaAKJanL/ssw0CXwyh0bYMOsLikBPI0+wPcBHZO9NvOz5/9ZXRkn0V2ui5OonAxrEigJ+nzfmH5QWeBN4JrMrYoOkif9dZTSN7m6TWAe8Cnhgv2VfhuF+TsAiUxlsCrE+j84+RF7if7PQLDJC/uxnayc4mn3hyepeZrUt6+3fCIrCDsgOARhOAYhqdf5hBduJrpp/JiOP00Zj6b5rYQzYOox0Ebg9jfnsdx3xSIrCu/D8KoyQvamnHXVlv5y8zyAC+hfgc6l83bcS6sSdP3h+e3+YW/zX6wxj8qJntaYDNkhCBteURymj151U1OP/7GuH8w/IC9wPH09r7CJ6qNWHUovyQ1u3x2I2vTv1rSFI3aszXKgI/GjMCCMpwVxXLgIY7/7C8wHP45MtTLTigdgPfJp+spDVPpN6Gr9Pfm1Syr0EisA3fWDduBLAa33iTeucfZpSt4aHcTms1DUUkl3xtNTaE8dZKrMUnoZ9qhvPXIAJFfL/JkJzZ1BGjMYpwzj2LP4b5r1rB+cs++++dc08BvwcOBl6b8sHUB3zMzH6VR++PoqjknOsLDpX2ikAJ+D5wqpm9EEVRGuz3a+fcKuAdwOsncP4vAddHUfTKRBEAZtaH35Tw9TFm02KYuVLj/GWffRDfgvkB0n/Y6cNBRPPMU2EspZnBEBW/P+xYTdN4HysSKOKTlNvx16cvGy1imTLBD28HjgmzaRvwW3x/8i5gddqcf5R8xmL8efJvTek68igz25RzAcDMHPAI6bz0oz9MKNeGykVabbiQoacoDYYIc/N4ojUlB4NrLnAj/lqptISZA/jS0e2IeFPQdcDHSNdhILuBZcDdWa3SFHIwuHqA0/B7CdKg4CXgQeBeuf4fn1EJvyU4TQnBHnxP/51ZLtFOzcMAi6LoD865Z4DfAH9Hc4+afhY438x+J9cf8oz+1zm3EX/td3eTP8464BTgh2b2SpbtPiVPgyyEmkvxR0434ziuDfhzETbL5cd8RqeEJVszNkeV8AnJc0JZOfNMzdPgiqLolSiKtjrnnsbvRZ/bQBHsAc42s7Vy87Fxzv0C+B3+mvVGnhNQBB4Azq1nT78EIB1CsMs5911gL2C/BtihBzjPzH4gF59YpJ1z64D/A95GYxK38ek9nzCz3jzZe2qOB1p/aBrqp75NQ5vxx0DL+Sf/bErOuZ+G/zwYeE0d364Xn+n/gpn9IW+2nprzgVZ0zv0YeAGfHHx9wm+xHjjNzFbLrSt+Ni87536EL8UdTH0Stz3AucADZvZyHu08VQMtesU591/4jrwD8EdU1ZoXKOI73M6KLz8RVUcCP8Pv9DyEZBuF1uLLw89mPdM/HlM0zF4l3N57HXAS1SegevEbkq7J23qyzs/mrfjqwBJq618pAo8DF6WtrVcRQPNnnD7n3JMhL3Agld1SU8Q3slwA3GZmv5dFE302LzrnHgv5gLdQ3Q1C/fgS8MVpbmNXBND82aaAv7Tj8/hy4USOvykMrAc169f92bThS4RXhGhgspHajvA9d+f08BUJQBWDbW4YNCcwtDutiN8QtQG4D38KrGaUxj6bbvz+jvPwm726xxHo5/CZ/ueTPrBTApCPGeeQMOu8HngxzPibga2aTZr+fDqDABwXntNsXq0Y7ATuAL6qyEwCIPIh1rN5tVrQI8cXQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEEEIIIYQQQgghhBBCCCGEECIVTJEJxsbM2oB9gUOBvwUGgBeB1cAGM9spK6XmWRWAmeE1CGw3sz5ZRgJQ7YCaCywDTgKmA4Wyv+4HtgOPAXcBG82sJKs17TkdBhwXxLobKAI7gPuAu81slywlAahkJlkKrAgDqjDBt+wEngBuANaZmYxY/2cEsA9wBnACMG+M51QCngcuB56RSI9kqkwwIuT/AHBLGFSTEcg/BRYBxwJtzrlfRFE0IGvW7RnNAa4CPgu8I0RnU8aZ4N4Ynk27c+5nURQNyoqKAEYbWF1hprgQ6KzyxxSBVcBlZrZKVk30+XSE2f4yYMEkIrPRns3jwCVmtk0WlQCUD65ZIeT/J6AtgR+5I4jJvWZWlIUTeT7X4vMxHTX+uPXAR4HntCTQEgAzWwDcCRydoD1eh09MdTrn1ijsrOn5LAbuAY4CXpPAj/xz4J3Ab51zP4+iKNciMCXHAwvgEODLwMI6vU0RuB/4qJntkTtX/IyWArcB8+vw4/uBW4Gr8vxspuZ0YBXwZaOvAW+u41sVgP2ANznnnomi6A9y60k/oyOB2/HJ2HrwGuBA4G+cc6ujKOpVBJCPgdUOXIDPJHc36G0VCVT2jA4BHgDmNugt1wEfAlbnrYw7NWcDqwv4DPBJfPmuUcSRwOtCJPCy3HzMZ7Q3cDc+098oZgHHAL92zm3KU15gao4G1kzgK8DpQHsTPkIBOADY5Zz7aRRF8vaRz2hGWJYd3ITotAs4EpjqnPtJFEW5qN4UcjKwFgDfAN5LMmW+aukIS48j5O4jnlEHcB2+C7NZ4zLuBbktTBiKAFp8UOGcOwTfE744JTmPTuAA59yjURRps0rAOXcmcCnJlPpq9YmFwJJQwv2fLNt9yiScqAOYE8LmPfEr7cmSkOl/D3AjMDtlH68EXA8sUzMKmNk+wPeoX8a/Wrbgk4MtsY8gtLJ3A9NCtNmH3xVZqlgAgoMvDWHZnBA6D+C3WvYBK8zszhQb4nzgGhqX6a+U3cDxZhbl3PnbgDvwuZm0PqfLgK+mtavTzDqBzwMuOH5H8NdB/Nb1S8ysZ9JLgOD8p+I75OaHtVFncKY/w3dTHeac+00URf+RMmN0hXX2FTQ2018pfwLMcs59I89VAefcUuBK4LUpfk6HA93OuR+nraszOP9NwFnAXmH2j/31dcBbgKOccz+IomjEtuixki1H4nfETZsgYXKDmZ2VImPMwHeOfYzae8YbwZKQm8jz7H9RiqO0mA7gYuAuM5udQuc/nfETp/sA3wp7KsaPAMxsOn7X1F6T+AztwD+mIRII9eN7SLanvxED62Xn3ON5LAs65w4N4fVrW+DjFsJs+jbn3Jow5lvB+WP2AmY75x4u/9yjfeOpVNaB1dRIwMwwsyXAI2EN1GqlzSNoXMdb2ji5BWb/4RwEfAc4JiSaW8H5Y04I0cDoAhDW/odX8ZmaIgJlmf5vUb8NPfVmNv5AkbyF/zPxSeZWZC6+tHxhaC1vBecHnxhcOl4OoAC8tcrP1lARCOvHC/Bn8s1uYV9ob2HxqoV98dWlVqUbXyFbYWbdLeD8MQeOJwCzwos0i0AwxLXA8hYMIUfjL3IqAJ0t/ju048vND5nZvBZwfoBF5T08w3/Q/gmsoesqAiF0vAufle3IiDN051AA3piR36OAr5p9x8yW1KNBLkHnJ+QA2scTgCSIReAjSRrEzPYFvk1yR3elhc05FICsbbZZGMbmKUkmB4Pz35KQ8xMmzfljCcB+CRqkK6yRLq5VBEKm/zB89nVJxgbOS/iNSnnjJ/jO0iwxE9+HclVw3KSc/1SSrW7tO0IAgpMmvQe7E9+OW7UIBDU9E3iI+hwN1exZcAWwMYcCEOFPUM4anfjzJu6ppWmojs4/ugAEFtTJIFWJQNiIdFUwxPQMDpYngJvzuCHIzPrx7dpZvLWnDb/1/LtmtriKcd9VR+cfEumX//BZdXSyWAT+ebLro7K23k+SnWRfOZuAj+f8/rrVYZmY1VOTF4Vl66kVjPtu/KEo9XJ+gL1HE4B6H8HUCVwNfGoiY4TtoY8Ap5CtZF9ML36H1pYcOz8h8vkS8CB+i3QWmRUmsuUT9QuENvx7QvRQzy7DveOopJECQJjJLwf+JTTyjFjvh9Ngv4e/kTeLJxYNBiF8CoGZDeD3A6zJ8K8Zbyb6VtizMpbz34c/m7De434aPmE55I3+uoHGuBy4pryNMqz3P4FP9s3L6EAo4U8HvlkHgQwZ/DuADwPbMvxrFvCXxXzfzI4uj4JDb8tD+H0hjZr09mZYeN3IDHt7UMQuM1sGzMBnw4/NaMgfsxK4NMx6Yihr8Vd23cP429BbnfnB2b9gZsvD2L8PX94uNPhzrCoEBWq0AMQicD7wfeBZmn9gZ73ZAnzYzHbK10eNAsBXRS4ne/0Bw+kCPgX8Wxj/hzRhuTt/+BKgGTX2Qljrz8/4A9+NvxRko1x9XBEo4Y9u/yLZ6xQcTluY9Rc06f3fXC4A08lnP3ojGMAfeaWk3+REoIhPkn6d7FYG0sC8cgGYJ3vUhXhG+4qSfhWJQB9wCb5bUEgAWpYngSvNTNeDVy4CO4HzgA2yRl2YbWbtsQDMkT0SZ31Y9/fKFFWLwBbgHGCHrFGXHMTsWAD+QvZIlJeA88xsq0xRM6vxPQIS0uSZowggeeL162qZIpEoAOAxfLeg+icSXgbEAjBDtkiEQfzmlm/m7Z75OotACbgd+BzZLw82khmxAEyTLWqmhN/U8rm0XiHV4iJQxO8ovR+VB5NieiwA6gGonZX47b0KU+snAgPAx4GnZY1EKMQC0CVb1MRm4ENmtkumqLsI7MLf2LtO1khAAcLXTpmianbjM/6bZIqGicBW4INAj6yRjAAoAqiOfmBZCP9FY1kbIgGVB2sRgHD4YEGmqJgicD1wt9p8mxIFgO+0XBaEWFQZAbTLDFXxKP5QE2X8mycCJeBOVB6seQkgKmMNcFE42VY0VwSK+Gvi7kblwaoEoE1mqIge4JxwjJVIhwjE5UFtua5CAJQAnDx78MdWrZcpUicCvfiNQ8/LGloC1INB/CUlj6vNN7UisD2IgMqDigASpRTWmF9Sxj/1rMOfI6DyoHIAiREBy3SwR0tEAYRcgHYPagmQCJvxnX6aUVpHBOLdg9ej8uCEAqA+gLHZje/x3yJTtJwIFPE5Gx0uOoEAaB/A6AziT/ONZIqWFYEB4CLgOVlDS4BKiDvMblXSr+VFYBe+MqAoTgIwaZ7DX+Gl9WM2RGAzqgxIACbJNpT0yyLPAJeiyoAEYBz68Ed5a29/9qIA8JWBm1FScIgAqA/AUwSW4y+oFNkUgRJwhZ7xUAFQFcDzJPBZJf0yLwL9+HsGFOVpCfBHtuK392p9mA8R6MGXB/skAKKIP9hDt/jki6fxh7pIAGQD/sHM1BGZLxYAh2rwiwJwErDczJQQzccSYA7wALoVWwIQaAPOB642M9kk284/E3gIWCRrSADKaQcuBj4lEcis808H7gOWyBoSgNHoAC6XCGTS+buBO4ClGvdDBUDGGCkClwEX6uivzDh/J3ALcKzG+0gBUOPLSDrxN9FeLBHIjPOfJOfXEkAikC/nbwduAk7VWJcASATy5/wr5PwSgKRE4FyZouWc/1y02U0CkJAIfN7MzpIp5PwSgHzSBdwgEZDzSwAkAhKB9Dl/J3CjnF8CIBHIp/PfBJwt569cANQHIBHIgvOfrgmtOgHokxkkAnJ+LQGEREDOLwEQEgE5f54EYI/MkJgIHClTNITlcv7kBEBJwOREYKHM0BAWyfklAGnkDTKB0BIgv8yQCUSrCYAuwEwONaE0BkWtCQqA+gCSQxeKNoYemSDZJYCigNrpA74hMzSEf9eYTUgAwl14uietNor4MwMimaIhPChbJ8JAXErRzFWDEfH3zn9WpwY1BjPbA5yGv95LVM+ThTJFHZQ9KmYDcCLwBd0q3HAReAl4dxDfnbJIxawF1k8pM+gngOtklzHpBzYC64GfAauB9WamtWjzxaATf9nHYuBAYH9gPqrKjBe1HmxmQwSggL844cwcG6YE7AY2h9cL+PzIJvwV4oMK81MvBvEfu4B98JeAviV8XQDsHf4uzxPZeWZ2L8CUUYx3CnAtMDejBujDl5Hi138H594KbAF2y8kzLQ4FYHaIEOYDbw5jPX7Nxl8Ok8XJbVVw/o3x/5wyhqEKwEH4a5RaNYzaE5x9T5jVdwI9ZqZavRhPJArAzCAGM4BpQHd4taow9AKPmtmI/on/B6/cY3m0e70ZAAAAAElFTkSuQmCC'
$FrontEndGraphic = 'data:image/png;base64,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'
$ServicesGraphic = 'data:image/png;base64,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'
$SQLGraphic = 'data:image/png;base64,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'
$InfoPaneGraphic = 'data:image/png;base64,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'
#endregion



#region functions
function Check-ServerCSService {
  [cmdletbinding()]
  Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $Server
  )
  Process
  {
    $NoError = $true
    if (Test-Connection -computer $Server -quiet -count 2) {
      $csw = get-CsWindowsService -computername $Server -ExcludeActivityLevel

      foreach ($svc in $csw){
        if ($svc.status -notlike 'Running') { $NoError = $false }

      }
    }else{ $NoError = $false }  # End If
    $TableReturn = $csw | Select Name,Status | ConvertTo-Html -Fragment
    $Return = @{
      'Result' = $NoError;
      'Services' = $TableReturn
    }
    return $Return
  } #Close Process
} #Close Function

function Check-ServerResponds {
  [cmdletbinding()]
  Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $Server
  )
  Process
  {
    $NoError = $false
    if (Test-Connection -computer $Server -quiet) { $NoError = $true }
      return $NoError
  } #Close Process
} #Close Function

function Build-StatusTable($ServerList,$pingonly) {
  $i=0
  Foreach ($PoolIdentity in $ServerList.Identity) {
    $i++
    Write-Debug -Message "Checking Server Count $($ServerList.count) - Pass $i"
    $percPool=$ServerList.identity.count
    Write-Progress -ParentId 1 -Activity "Getting Pool Status" -Status "Querying Pool: $PoolIdentity" -Id 2 -PercentComplete $($i/$percPool*100)
    $ArrComputer =@() # Create a blank Array For Shortened Computer NETBIOS name
    $ArrFQDNComputer =@() # Create a blank Array For FQDN name
    $FilteredPools = $ServerList | Where-Object { $_.identity -eq $PoolIdentity} #create a new object for each 'group' or servers within a pool
    $j=0
    Foreach ($Computer in $FilteredPools.Computers) {
      
      Write-Debug "Got to server recussion fileterpools count $($filteredpools.computers.count) - pass $j"
      # Loop through the computers in each pool
      $strComputer = $Computer.ToString()
      $percServer=$FilteredPools.computers.count
      $ArrFQDNComputer += $strComputer
      # Remove Trailing Domain name if there - everything after '.'
      $strComputer = $strComputer.Substring(0,$strComputer.IndexOf('.'))
      #add the filtered computer name to the array
      $ArrComputer += $strComputer
      
    }
    Write-verbose $poolIdentity 
    $CommonServerName = get-LongestCommonSubstringArray $ArrComputer
    $ReturnHTML += "<h3>$poolIdentity ( $CommonServerName )</h3><table><tr>"
    $arrFQDNcomputer | ForEach-Object { 
      write-verbose -message  "Checking Services on $_"
      $j++
      Write-Progress -ParentId 2 -Activity "Getting Server Status" -Status "Querying $strComputer within the $PoolIdentity pool " -Id 3 -PercentComplete $($j/$percServer*100)
      #
      # Lets check to see if we have already scanned this machine
      #

      if ($HashCheckedComputers.ContainsKey($_)) {
        $ServerServiceOK = $HashCheckedComputers.get_item($_)
        If ($ServerServiceOK.result) { 
          $ComputerTDclass = 'computer_pass'
        } else {
          $ComputerTDclass = 'computer_fail'
          $global:ErrorCount++
        }
      } else {
        # Not in prescanned list - lets check the server is ok
        if ($pingonly) {
          $ServerOK = $_ | Check-ServerResponds
          If ($ServerOK) { 
            $ComputerTDclass = 'computer_pass'
          } else {
            $ComputerTDclass = 'computer_fail'
            $global:ErrorCount++
          }
        } else {
          $ServerServiceOK = $_ | Check-ServerCSService
          If ($ServerServiceOK.result) { 
            $ComputerTDclass = 'computer_pass'
          } else {
            $ComputerTDclass = 'computer_fail'
            $global:ErrorCount++
          }
          $HashCheckedComputers.Add($_, @{
            result = $ServerServiceOK.result
            services = $ServerServiceOK.services
          }) # We use this so that the program does not check the same server multiple times
        }
        
      }
      $strHTMLComputer = $_.Substring(0,$_.IndexOf('.'))
      if ($CommonServerName -ne $null){ 
        if ($overrideAutoStrip){
          $strHTMLComputer = $_.Substring(0,$_.IndexOf('.'))
          $strHTMLComputer = $strHTMLComputer.substring($strHTMLComputer.length - $charstokeepforserverid,$charstokeepforserverid)
        } else {
           $strHTMLComputer = ($_.Substring(0,$_.IndexOf('.'))) -replace "$CommonServerName(.*)",'$1'  
        }
      }

      $ReturnHTML += "<td class='$ComputerTDclass'> <div class='tooltip'>$strHTMLComputer <span class='tooltiptext'>$($ServerServiceOK.services)</span></div></td>"
    } 
    $ReturnHTML += '</tr></table>'
    $ArrComputer = $null


  } # end foreach
  return $ReturnHTML
} # end function


function Check-Ping($server)
{
  if (Test-Connection -ComputerName $server -Quiet) {$pingresult = $true} else {$pingresult =$false}
  return $pingresult
}


 function Get-PoolIPAddress($server)
{
  if (Test-Connection -ComputerName $server -Quiet) {$pingresult = $true} else {$pingresult =$false}
  return $pingresult
}


function Check-ReplicationStatus($ReplicaFqdn)
{
  $ReplicaStatus = (Get-CsManagementStoreReplicationStatus -ReplicaFqdn $ReplicaFqdn).uptodate
  return $ReplicaStatus
}

Function Convert-HTMLEscape 
{
  <#
      Convert &lt; and &gt; to < and >
      It is assumed that these will be in pairs
      Also convert $quot; to "
  #>
  [cmdletbinding()]
  Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    [string[]]$Text
  )
  Process
  {
    foreach ($item in $Text)
    {
      if ($item -match '&lt;BREAK-NL&gt;')
      {
        $item.Replace('&lt;BREAK-NL&gt;','<BR>')
      }
      else
      {
        #otherwise just write the line to the pipeline
        $item
      }
    }
  } #close process
} #close function


Function get-LongestCommonSubstringArray
{
  Param(
    [Parameter(Position=0, Mandatory=$True)][Array]$Array
  )
  $PreviousSubString = $Null
  $LongestCommonSubstring = $Null
  foreach($SubString in $Array)
  {
    if($LongestCommonSubstring)
    {
      $LongestCommonSubstring = get-LongestCommonSubstring $SubString $LongestCommonSubstring
      write-verbose "Consequtive diff: $LongestCommonSubstring"
    }else{
      if($PreviousSubString)
      {
        $LongestCommonSubstring = get-LongestCommonSubstring $SubString $PreviousSubString
        write-verbose "first one diff: $LongestCommonSubstring"
      }else{
        $PreviousSubString = $SubString
        write-verbose "No PreviousSubstring yet, setting it to: $PreviousSubString"
      }
    }
  }
  Return $LongestCommonSubstring
}

Function Add-HTMLTableAttribute
{
  Param
  (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [string]
    $HTML,

    [Parameter(Mandatory=$true)]
    [string]
    $AttributeName,

    [Parameter(Mandatory=$true)]
    [string]
    $Value

  )

  $xml=[xml]$HTML
  $attr=$xml.CreateAttribute($AttributeName)
  $attr.Value=$Value
  $xml.table.Attributes.Append($attr) | Out-Null
  Return ($xml.OuterXML | out-string)
}    

Function get-LongestCommonSubstring
{
  [cmdletbinding()]
  Param(
    [string]$String1, 
    [string]$String2
  )

  if((!$String1) -or (!$String2)){Break}


  # .Net Two dimensional Array:
  $Num = New-Object 'object[,]' $String1.Length, $String2.Length
  [int]$maxlen = 0
  [int]$lastSubsBegin = 0
  $sequenceBuilder = New-Object -TypeName 'System.Text.StringBuilder'

  for ([int]$i = 0; $i -lt $String1.Length; $i++)
  {
    for ([int]$j = 0; $j -lt $String2.Length; $j++)
    {
      if ($String1[$i] -ne $String2[$j])
      {
        $Num[$i, $j] = 0
      }else{
        if (($i -eq 0) -or ($j -eq 0))
        {
          $Num[$i, $j] = 1
        }else{
          $Num[$i, $j] = 1 + $Num[($i - 1), ($j - 1)]
        }
        if ($Num[$i, $j] -gt $maxlen)
        {
          $maxlen = $Num[$i, $j]
          [int]$thisSubsBegin = $i - $Num[$i, $j] + 1
          if($lastSubsBegin -eq $thisSubsBegin)
          {#if the current LCS is the same as the last time this block ran
            [void]$sequenceBuilder.Append($String1[$i]);
          }else{ #this block resets the string builder if a different LCS is found
            $lastSubsBegin = $thisSubsBegin
            $sequenceBuilder.Length = 0 #clear it
            [void]$sequenceBuilder.Append($String1.Substring($lastSubsBegin, (($i + 1) - $lastSubsBegin)))
          }
        }
      }
    }
  }
  return $sequenceBuilder.ToString()
}


function Check-ServicesAreRunning{
  [cmdletbinding()]
  Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $Server
  )
  Process
  {
    $CheckServer = Get-CsWindowsService -ComputerName $Server -ExcludeActivityLevel
    $anError = $True
    foreach ($svc in $CheckServer)
    {
      if ($svc.status -notlike 'Running') {$anError = $false}
    }
    $anError
  } #close process
} #close function


function Check-CSDatabaseServicesAreRunning{
  [cmdletbinding()]
  Param (
    [Parameter(Position=0,ValueFromPipeline=$True)]
    $PoolIdentity
  )
  Process
  {
    $IdentityToCheck = "UserDatabase:$PoolIdentity"
    $NoError = $False
    if (Test-Connection -ComputerName $PoolIdentity -Quiet)
    {
      if ((Get-CsUserDatabaseState -Identity $IdentityToCheck).online -notlike 'True') {$NoError = $False} else {$NoError = $True}
    }
    return $NoError
  } #close process
} #close function

Function Build-DatabaseTable()
{
  $ReturnHTML = '<table><tr><th>SQL Pool</th><th>Online</th></tr>'
  foreach ($dbpool in $objDatabaseServers.identity){
    $ServerOK = $dbpool | Check-CSDatabaseServicesAreRunning
    If ($ServerOK) { 
      $ComputerTDclass = 'computer_pass'
    } else {
      $ComputerTDclass = 'computer_fail'
      $global:ErrorCount++
    }
    $strHTMLComputer = $dbpool.Substring(0,$dbpool.IndexOf('.'))
    $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$ServerOk</td></tr>"
  } # end for each
  $ReturnHTML += '</table>'
  return $ReturnHTML
} #end function


function Build-CSIMTable(){
  $ReturnHTML = '<table><tr><th>Pool</th><th>Completed</th><th>Latency</th></tr>'
  $PoolsWithTestAccounts = Get-CsHealthMonitoringConfiguration
  foreach ($dbpool in $PoolsWithTestAccounts.identity){
    $ServerOK = Test-CsIM -TargetFqdn $dbpool
    If ($ServerOK.Result -like 'Success') { 
      $ComputerTDclass = 'computer_pass'
    } else {
      $ComputerTDclass = 'computer_fail'
      $global:ErrorCount++
    }

    $strHTMLComputer = $dbpool
    $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$($ServerOk.result)</td><td>$($ServerOk.Latency)</td></tr>"
  } # end for each
  $ReturnHTML += '</table>'
  return $ReturnHTML

}


function Build-CSAddressBookTable(){
  $ReturnHTML = '<table><tr><th>Pool</th><th>Completed</th><th>URI</th></tr>'
  $PoolsWithTestAccounts = Get-CsHealthMonitoringConfiguration
  foreach ($dbpool in $PoolsWithTestAccounts.identity){
    $ServerOK = Test-CsAddressBookService -TargetFqdn $dbpool
    If ($ServerOK.Result -like 'Success') { 
      $ComputerTDclass = 'computer_pass'
    } else {
      $ComputerTDclass = 'computer_fail'
      $global:ErrorCount++
    }

    $strHTMLComputer = $dbpool
    $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$($ServerOk.result)</td><td><div style='word-break:break-all;'>$($ServerOk.TargetURI)</div></td></tr>"
  } # end for each
  $ReturnHTML += '</table>'
  return $ReturnHTML

}

function Build-FederationPartnerTable(){
  $ReturnHTML = '<table><tr><th>Partner</th><th>Success</th></tr>'
  
  # Lets Find Out Which Pool is responsible for federation
  $FedService = (Get-CsTopology).sites.siteconfiguration.federationroute.localname
  $FedServer = ((Get-CsService -EdgeServer) | where {$_.serviceid -eq $FedService}).Poolfqdn
  
  foreach ($partner in $FederationPartners){
    $ServerOK = Test-CsFederatedPartner -Domain $partner -TargetFqdn $FedServer
    If ($ServerOK.Result -like 'Success') { 
      $ComputerTDclass = 'computer_pass'
    } else {
      $ComputerTDclass = 'computer_fail'
      $global:ErrorCount++
    }

    $strHTMLComputer = $partner
    $ReturnHTML += "<tr><td>$strHTMLComputer</td><td class='$ComputerTDclass'>$($ServerOk.result)</td></tr>"
  } # end for each
  $ReturnHTML += '</table>'
  return $ReturnHTML

}

#endregion


#region Main Program


#region Verifying Administrator Elevation 
write-verbose -message  "Verifying User permissions..."
#Start-Sleep -Seconds 1 
#Verify if the Script is running under Admin privileges 
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')){
    Write-Warning "You do not have Administrator rights to run this script.`nPlease re-run this script as an Administrator!" 
    Break
}
else 
{ 
  write-verbose -message  ' Done!' 
} 
#endregion 

#region Import Lync / SfB Module 

write-verbose -message  "Please wait while we're loading Skype for Business / Lync PowerShell Modules..."  
if(-not ((Get-Module -Name 'Lync')) -or (Get-Module -Name 'SkypeforBusiness')){ 
  if(Get-Module -Name 'Lync' -ListAvailable){ 
    Import-Module -Name 'Lync'

  } 
  elseif (Get-Module -Name 'SkypeforBusiness' -ListAvailable){ 
    Import-Module -Name 'SkypeforBusiness'

  } 
  else{ 
    throw  "Lync/SfB Modules do not exist on this computer, please verify the Lync Admin tools installed  "
    

  }     
} 

write-verbose -message  ' Done!' 
#endregion 


#region Collating pools and servers
$Pools = Get-CsPool | Select-Object Identity,Services | Sort-Object -Property Identity
[int]$global:ErrorCount = 0

write-verbose -message  'Finding All Pools & Services in environment..' 
Foreach ($pool in $Pools)
{

  Write-verbose "[ ] $($Pool.Identity)" 
  $PoolServices = $pool.Services
  Foreach ($service in $PoolServices)
  {
    Write-verbose " - [ ] $service" 

  }
}

write-verbose -message  ' Done!' 



$objRegistrarServers = Get-CsPool | Where-Object {$_.Services -like 'Registrar:*'} | Select-Object identity,computers
$objMediationServers = Get-CsPool | Where-Object {$_.Services -like 'MediationServer:*'} | Select-Object identity,computers
$objEdgeServers = Get-CsPool | Where-Object {$_.Services -like 'EdgeServer:*'} | Select-Object identity,computers
$objDatabaseServers = Get-CsPool | Where-Object {$_.Services -like 'UserDatabase:*'} | Select-Object identity,computers


#endregion

#region Building the HTML
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Serices List" -Id 1 -PercentComplete 10

$strPoolsServicesHTML = '<div id="Overview" Class="ContentDiv"><h1>Overview</h1><div class="ContentInnerDiv">'
$strPoolsServicesHTML += "<H3>Pools,Services and Versions</h3><h5 class='subtitle'>Services and pools within each pool, use this table to identify what each pools role is</h5>"

$strPoolsServicesHTML += "<div id='pooldist' class='pooldist'><h2 class='divheader'>Pool Distribution</h2><div class='PoolDistInnerDiv'>"
$strPoolsServicesHTML += "<img class='poolgraphic' src='$(Get-PoolDoughnut)' />"

$strPoolsServicesHTML += "</div></div>"

$UserCounts = Get-CBUserCounts
$strPoolsServicesHTML += @"
          <div class="usercountinfo" id="usercountinfo">
                    <h2 class="divheader">User Counts</h2>
                    <div class="PoolDistInnerDiv">
                        <div class="UserCounts">
                            <div class="totalusers" id="Ttlusers"> 
                                <h3>Total Users</h3>
                                <div id="totalusersspan" class="totalusersspan">$($UserCounts.AllUsers)</div>
                                </div>
                            <div class="totalusers" id="EVusers"> 
                                <h3>EV Users</h3>
                                <div id="totalusersspan" class="totalusersspan">$($UserCounts.EVUsers)</div>
                            </div>
                              <div class="totalusers" id="ExUMusers"> 
                                <h3>Exchange UM Users</h3>
                                <div id="totalusersspan" class="totalusersspan">$($UserCounts.ExUMUsers)</div>
                                
                            </div>
                        </div>
                    </div>
                </div>
"@





$strPoolsServicesHTML += "<div id='poolinfo' class='poolinfo'><h2 class='divheader'>Pool Information</h2><div class='pooltable'>"
$PoolsTableWithAttib = ($replaceBreak = $pools | Select-Object Identity,@{Name='Services'
Expression={[string]::Join('<BREAK-NL>',($_.Services))}} | ConvertTo-Html -Fragment | Out-String | Add-HTMLTableAttribute -AttributeName 'class' -Value 'pools').replace('&lt;BREAK-NL&gt;','<br>')
$strPoolsServicesHTML += $PoolsTableWithAttib
$strPoolsServicesHTML += '</div></div>'


$strPoolsServicesHTML += "<div class='poolinfo'><h2 class='divheader'>Version Information</h2><div class='pooltable'>"
$strPoolsServicesHTML += Get-CsManagementStoreReplicationStatus | Select-Object ReplicaFqdn,ProductVersion | ConvertTo-Html -Fragment

$strPoolsServicesHTML += '</div></div></div>'

$strPoolsServicesHTML += '</div>' # Closing Overview pane


write-verbose -message  'Checking Servers in the environment..' 



$strCsWindowsServiceHTML = '<div id="FrontEnd" Class="ContentDiv"><h1>Server Checks</h1><div class="ContentInnerDiv">'


$strCsWindowsServiceHTML += '<H3>Individual Server Check</h3>'
$strCsWindowsServiceHTML += "<h5 class='subtitle'>Dashboad showing each server sorted by type then pool, use this table to identify if an individual server is running all windows services and/or responding to pings</h5>"

$HashCheckedComputers =@{}
#Checking Registar Servers
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Registrar Servers" -Id 1 -PercentComplete 20
$strCsWindowsServiceHTML += "<div class='FEandMed_Container'>"
$strCsWindowsServiceHTML += "<div class='servers'><h2 class='divheader'>Registrar Servers</h2><div class='serverdivcontent'>"
$strCsWindowsServiceHTML += Build-StatusTable $objRegistrarServers $false
$strCsWindowsServiceHTML += '</div></div>'

#Checking Mediation Servers
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Mediation Servers" -Id 1 -PercentComplete 30
$strCsWindowsServiceHTML += "<div class='servers'><h2 class='divheader'>Mediation Servers</h2><div class='serverdivcontent'>"
$strCsWindowsServiceHTML += Build-StatusTable $objMediationServers $false
$strCsWindowsServiceHTML += '</div></div>'
$strCsWindowsServiceHTML += '</div>'

#Checking Edge Servers
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Edge Servers" -Id 1 -PercentComplete 40
$strCsWindowsServiceHTML += "<div class='EdgeandSQL_Container'>"
$strCsWindowsServiceHTML += "<div class='servers'><h2 class='divheader'>Edge Servers</h2><div class='serverdivcontent'>"
$strCsWindowsServiceHTML += Build-StatusTable $objEdgeServers $true 
$strCsWindowsServiceHTML += '</div></div>'

#Checking SQL Servers Respond to Ping
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - SQL Servers" -Id 1 -PercentComplete 50
$strCsWindowsServiceHTML += "<div class='servers'><h2 class='divheader'>Database Servers</h2><div class='serverdivcontent'>"
$strCsWindowsServiceHTML += Build-StatusTable $objDatabaseServers $true 
$strCsWindowsServiceHTML += '</div></div>' #closing the section
$strCsWindowsServiceHTML += '</div></div>' # Glosing the dual layout
$strCsWindowsServiceHTML += '</div>' #Closing the page div 'frontend'


#Lets the Hash table
$HashCheckedComputers.Clear()


write-verbose -message  ' Done!'
write-verbose -message  'Checking Services in the environment..'
Write-Progress -id 3 -Completed -Activity "Completed"
Write-Progress -id 2 -Completed -Activity "Completed"
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Replication Services" -Id 1 -PercentComplete 60

$strReplicationHTML = '<div id="SQLServer" Class="ContentDiv"><h1>SQL Level</h1><div class="ContentInnerDiv">'

$strReplicationHTML += '<H3>Lync Database Check</h3>'
$strReplicationHTML += "<h5 class='subtitle'>Dashboad showing the status of replication and database services</h5>"

$strReplicationHTML += "<div class='Replication_Container'>"

$strReplicationHTML += "<div class='repservices'><h2 class='divheader'>CMS Management Replication Status</h2><div class='serverdivcontent'>"
$strReplicationHTML += Get-CsManagementStoreReplicationStatus | Select-Object @{Name='Replica';Expression={($_.ReplicaFqdn.Substring(0,$_.ReplicaFqdn.IndexOf('.')))}},uptodate,LastStatusReport,LastUpdateCreation | ConvertTo-Html -Fragment
$ErrorMatches = Select-String -InputObject $strReplicationHTML -Pattern '<td>False</td>' -AllMatches
$global:ErrorCount = $global:ErrorCount + $ErrorMatches.Matches.Count
$strReplicationHTML = $strReplicationHTML.replace('<td>True</td>',"<td class='replica_pass'>True</td>")
$strReplicationHTML = $strReplicationHTML.replace('<td>False</td>',"<td class='replica_fail'>False</td>")
$strReplicationHTML += '</div></div>'

Write-Progress -Activity "Generating Report" -Status "Constructing HTML - User Database Services" -Id 1 -PercentComplete 70
$strReplicationHTML += "<div class='sqlservices'><h2 class='divheader'>User Database Online</h2><div class='serverdivcontent'>"
$strReplicationHTML += Build-DatabaseTable
$strReplicationHTML += '</div></div>'

$strReplicationHTML += '</div></div>'
$strReplicationHTML += '</div>' # closing SQL Div section


write-verbose -message  ' Done!'


write-verbose -message  'Checking Communications within the environment..'
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Communications Services" -Id 1 -PercentComplete 80
$strHealthHTML = '<div id="Services" Class="ContentDiv"><h1>Lync Services</h1><div class="ContentInnerDiv">'
$strHealthHTML += '<H3>Lync Communications Check</h3>'
$strHealthHTML += "<h5 class='subtitle'>Are communications between the two test accounts working?</h5>"

$strHealthHTML += "<div class='connectivity_Container'>" # open container

$strHealthHTML += "<div class='servers'><h2 class='divheader'>Health Monitoring Config</h2><div class='serverdivcontent'>"
$strHealthHTML += Get-CsHealthMonitoringConfiguration | Select-Object Identity,FirstTestUserSipUri,FirstTestSamAccountName,SecondTestUserSipUri,SecondTestSamAccountName,TargetFqdn  | ConvertTo-Html -Fragment -as List
$strHealthHTML += '</div></div>'

$strHealthHTML += "<div class='servers'><h2 class='divheader'>IM between test users</h2><div class='serverdivcontent'>"
$strHealthHTML += Build-CSIMTable
$strHealthHTML += '</div></div>'
$strHealthHTML += '</div>' # close container


write-verbose -message  'Checking Address Book Retrieval within the environment..'
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Address Book Services" -Id 1 -PercentComplete 90
$strHealthHTML += '<H3>Address Book Check</h3>'
$strHealthHTML += "<h5 class='subtitle'>Tests the Address Book Download Web service by using test users preconfigured for the pool</h5>"

$strHealthHTML += "<div class='addressbook_Container'>" # open container

$strHealthHTML += "<div class='addressservices'><h2 class='divheader'>Health Monitoring Config</h2><div class='serverdivcontent'>"
$strHealthHTML += Build-CSAddressBookTable
$strHealthHTML += '</div></div>'


write-verbose -message  ' Done!'

write-verbose -message  'Checking Partner Federation..'
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Federation Partner" -Id 1 -PercentComplete 92



$strHealthHTML += "<div class='federationservices'><h2 class='divheader'>Partner Checks</h2><div class='serverdivcontent'>"
$strHealthHTML += Build-FederationPartnerTable
$strHealthHTML += '</div></div>'


$strHealthHTML += '</div></div>' # close container
$strHealthHTML += '</div>' # closing page div 'services'
write-verbose -message  ' Done!'



write-verbose -message  'Building HTML Report..'
Write-Progress -Activity "Generating Report" -Status "Constructing HTML - Finalizing" -Id 1 -PercentComplete 95

#get-LongestCommonSubstring "BigMan-LyncFE1" "BigMan-LyncFE2"

$strDivandImgHTML = @"

    <div id="logo" class="logo">

                    <img class="logo" src="$CompanyLogoBase64"/>

    </div>
    <div class='mainframe'>
    <div id="Nav" class="nav">
      <div id="navbuts" class="navbuts">

        <div id="navbut1" class="navbut" onclick="InfoPanes('Overview','navbut1')">
          <div class="menuword">Overview</div><div class="menuicon">
            <img class="menuimg" src="$OverviewGraphic"/>
          </div>
        </div>

        <div id="navbut2" class="navbut" onclick="InfoPanes('Capacity','navbut2')">
          <div class="menuword">Capacity</div>
          <div class="menuicon"><img class="menuimg" src="$CapacityGraphic"/></div>
        </div>

        <div id="navbut3" class="navbut" onclick="InfoPanes('FrontEnd','navbut3')">
          <div class="menuword">Servers Status</div>
          <div class="menuicon"><img class="menuimg" src="$FrontEndGraphic"/></div>
        </div>

        <div id="navbut4" class="navbut" onclick="InfoPanes('Services','navbut4')">
          <div class="menuword">Lync Services</div>
          <div class="menuicon"><img class="menuimg" src="$ServicesGraphic"/></div>
        </div>

        <div id="navbut5" class="navbut" onclick="InfoPanes('SQLServer','navbut5')">
          <div class="menuword">SQL Servers</div>
          <div class="menuicon"><img class="menuimg" src="$SQLGraphic"/></div>
        </div>

        <div id="navbut6" class="navbut" onclick="InfoPanes('ALL','navbut6')">
          <div class="menuword">ALL PANELS</div>
          <div class="menuicon"><img class="menuimg" src="$InfoPaneGraphic"/></div>
        </div>
      </div>
    </div>
    <div id="content" class="content">

"@






$PostDIVContent = @"
    </div>
    </div>
            <div id="IssueCount" Class="IssueCount">
                <div id="IssueCountNumber" Class="IssueCountNumber">$global:ErrorCount</div>
                <div id="IssuesTitle" class="issuesTitle">Issue(s) found</div>
        </div>
        <div id="reportdate" class="reportdate">
          Reported Created $(get-date) on $env:computername
        </div>

"@
#ConvertTo-Html -Head $Header,$JavaScript -Body $strDivandImgHTML,$strCsWindowsServiceHTML -PostContent $PostDIVContent | Out-File c:\PShtml.html
ConvertTo-Html -Head $Header,$JavaScript -Body $strDivandImgHTML,$strPoolsServicesHTML,$strCsWindowsServiceHTML,$strReplicationHTML,$strHealthHTML -PostContent $PostDIVContent | Out-File $home\Documents\PShtml.html
Invoke-Item $home\Documents\PShtml.html

#endregion

write-verbose -message  ' Done!'
#endregion

}



function Get-CBUserCounts {
    
    [cmdletbinding()]
     Param (
      
    )


$EntirePoolCount = Get-CsUser  # Lets
$PoolCount = $EntirePoolCount | Group-Object RegistrarPool | Select Count,Name | Sort-Object Count -Descending
$EVUsers = ($EntirePoolCount | where {$_.EnterpriseVoiceEnabled -eq $true}).count
$ExUMUsers = ($EntirePoolCount | where {$_.ExUMEnabled -eq $true}).count

$Return = @{
'AllUsers' = $EntirePoolCount.Count;
'PoolGroups' = $PoolCount;
'EVusers' = $EVUsers;
'ExUMusers' = $ExUMUsers;
}
Write-Output $Return

}


function Get-CBlyncBase64 {

    <#
      .SYNOPSIS
      Gets the Base64 of an image url that is passed to it.

      .DESCRIPTION
      The function is designed to convert a image URL to a base64 code. If you pass a url, the software will visit that site. subject to there being an internet connection 
      and download the image into it's parser, from there the code will convert the binary to a base 64 code

      .EXAMPLE
      Get-CBlyncBase64 -url http://www.google.com/logo.jpg

      .NOTES
      This function was really designed to assist the health HTML generation.

      .LINK
      https://www.linkedin.com/in/cburns/

      .INPUTS
      No Inputs needed

      .OUTPUTS
      
    #>
    [cmdletbinding()]
     Param (
       #Get the location of the image
      [Parameter(Mandatory)]
      [string]$Url
    )

    Process{
      if ($url -like "Http*"){
        $image = Invoke-WebRequest $url
        [string]$downloadedImgCheck = $image.Headers.'content-type'
          if ($downloadedImgCheck -like "image*") {
            $Base64toReturn = [convert]::ToBase64String($image.content)
          }
      }else{
          $image = Get-Content $Url -Encoding byte
          $Base64toReturn = [convert]::ToBase64String($image)
      }

    }
    END{
      Return "data:image/png;base64,$Base64toReturn"
    }

    

}

Function Get-PoolDoughnut {
    $PoolCount = (Get-CBUserCounts).poolgroups


    # load the appropriate assemblies 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")



    $PoolUsageChart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart

    $PoolUsageChart1.Width = 800

    $PoolUsageChart1.Height = 400

    $PoolUsageChart1.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")


    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea

    $chartarea.Name = "ChartArea1"
    $chartarea.BackColor =[System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $PoolUsageChart1.ChartAreas.Add($chartarea)
    
    [void]$PoolUsageChart1.Series.Add("data1")

    $PoolUsageChart1.Series["data1"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut

    $PoolList = @(Foreach($pool in $PoolCount.Name){$pool})
    $PoolNumber = @(Foreach($number in $PoolCount){$number.Count})

    $PoolUsageChart1.Series["data1"].Points.DataBindXY($PoolList, $PoolNumber)
    #$PoolUsageChart1.Series['data1'].Palette = "SeaGreen"

    #There is going to be an anyoing bug here... If there are more than 8 entries then you will have another label with black text on a black circle... DAMN!

    $PoolUsageChart1.Series['data1'].Points[0].LabelForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee") 
    $PoolUsageChart1.Palette = [System.Windows.Forms.DataVisualization.Charting.ChartColorPalette]::None
    $PoolUsageChart1.PaletteCustomColors = @([System.Drawing.Color]::Black,[System.Drawing.ColorTranslator]::FromHtml("#666666"),[System.Drawing.ColorTranslator]::FromHtml("#ffed00"), [System.Drawing.ColorTranslator]::FromHtml("#64ff00"), [System.Drawing.ColorTranslator]::FromHtml("#00c9ff"), [System.Drawing.ColorTranslator]::FromHtml("#d9d9d9") )

    # set chart options 

    $PoolUsageChart1.Series["data1"]["PieLineColor"] = "Black" 
    $PoolUsageChart1.Series["data1"].Font = "Arial,18"

    $PoolUsageChart1.Series["data1"].CustomProperties = "DoughnutRadius=50, PieLabelStyle=inside,PieDrawingStyle=default,PieStartAngle=270"
    $PoolUsageChart1.Series["data1"].IsValueShownAsLabel = $true

    $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.IsEquallySpacedItems = $True
    $Legend.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

    $PoolUsageChart1.Legends.Add($Legend)
    $PoolUsageChart1.Series[‘data1’].LegendText = "#VALX (#VALY)"

    $tempfile = New-TemporaryFile

    $PoolUsageChart1.SaveImage($tempfile,"png")



    Get-CBlyncBase64 $tempfile


}