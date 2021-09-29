#Debut de script
param ([switch]$Refresh )
 
# write by SLESIRE 
# Version 1.4
# modifié le 29/09/21
# ner pas lancer dans ISE
 
 
 
write-host "$(Get-Date) ### Init Script" -ForegroundColor Yellow
 
 
$ErrorActionPreferenceold = $ErrorActionPreference
$ErrorActionPreference = "SilentlyContinue"
 
 
if ($Refresh) {write-host "$(Get-Date) ### >> Mode Refresh Picture : ON" -ForegroundColor Green}
 
$scriptPath = [STRING]$pwd
#$passwordFilePath = $scriptpath + "D:\Annuaire\O365password.txt"
 
# Folder Path
  
$OutPath = "c:\Inetpub\wwwroot\annuaire"
#$OutPath = "D:\tmp\Annuaire"
$ImagesPath = $OutPath + "\Images\"
$LogPath = $OutPath + "\logs\"
$HTMLPath = $OutPath
$OutputPath = $OutPath
Write-Host $scriptPath
 
$MessgeTxtFile = $scriptPath + "\message.txt"

$messageTxt = get-content $MessgeTxtFile
 
$CSVFile = $LogPath + "Tall_user_O365.csv" 
 
$htmlFile = $HTMLPath + "\index.html"
 
 
if (-Not (test-path $OutPath) ) {
    $msg = "  --- >  /!\ folder $OutPath does not exist " 
    write-Host $msg -ForegroundColor Red
    write-host "Script is canceled"
    exit
    }
 
 
 
if (-Not (test-path $logpath) ) {
    write-host "$(Get-Date) ### >> Create log Folder" -ForegroundColor Yellow
 
    mkdir $LogPath| out-Null
    }
 
if (-Not (test-path $ImagesPath) ) {
    write-host "$(Get-Date) ### >> Create Images Folder" -ForegroundColor Yellow
 
    mkdir $ImagesPath | out-Null
    }
 
 
write-host "$(Get-Date) ### >> Connect Office 365 - Azure Active directory" -ForegroundColor Yellow
 
# connect to O365
 
# DEV
 
$User = 'ton compte'
$Password = 'ton mot de passe'
$O365Cred = $(new-object System.Management.Automation.PSCredential($User,(ConvertTo-SecureString -String $Password -AsPlainText -Force)))
$SessionExchO365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $o365cred -Authentication Basic -AllowRedirection -ErrorAction 'Stop' -ErrorVariable 'ConnectionError'  
# import Module Azure AD
Import-Module MSOnline
# connect to O365
#Connect-MsolService -Credential $o365cred
 
Connect-MsolService -Credential $O365Cred
write-host "$(Get-Date) ### >> Connect Office 365" -ForegroundColor Yellow
 
# connect to Exchange online　
write-host "$(Get-Date) ### >> Connect Exchange Online" -ForegroundColor Yellow
Import-PSSession $SessionExchO365 -AllowClobber -ErrorAction 'Stop' -ErrorVariable 'SessionError'    | out-Null
 
$AllUserObject = @()
 
write-host "$(Get-Date) ### >> Get All Users" -ForegroundColor Yellow
 

$allUser = Get-MsolUser -All -EnabledFilter EnabledOnly | where  { $_.userprincipalname -notlike "*#EXT#@*" -OR ($_.Department -notlike "exclude" )}
$AllGetUserObj = Get-User -ResultSize unlimited
$AllMailbox =Get-Mailbox -ResultSize unlimited | where {$_.Haspicture -eq $true}
write-host "$(Get-Date) ### >> Begin Analyze" -ForegroundColor Yellow
 
$i=1
$allUser | where {($_.Department -notlike "exclude" )}  |  % {
    # PROGRESS BAR Home Made ;-)
    $i++
    $percent = (($i / $allUser.Count)  * 100)
    Write-Progress -activity "Traitement en cours" -status "Percent added: $Percent %" -PercentComplete $Percent
 
    $Props =@{
        UserPrincipalName = $null
        IsLicensed = $null
        email = $null
        LastName = $null
        FirstName = $null
        MobilePhone = $null
        Phone = $null
        manager = $null
        manageremail = $null
        Identity = $null
        Office = $null
        Department = $null
        CountryOrRegion = $null
        AccountDisabled = $null
        ImageName = $null
        Title = $null
    }
    $userObjet = new-object psobject -Property $props 
          
    $Upn = [STRING]$_.UserPrincipalName
    $userObjet.UserPrincipalName = $Upn
    #$Upn 
    $isLicensedUser = $_.IsLicensed
    $userobjet.IsLicensed = $_.IsLicensed
 
    $GetUserObj = ''
    $GetUserObj = $AllGetUserObj | Where {$_.UserprincipalName -eq $UPN}
    if($GetUserObj)
        {
 
            $userObjet.LastName = $GetUserObj.LastName
            $userObjet.FirstName = $GetUserObj.FirstName
            $userObjet.email = $GetUserObj.WindowsEmailAddress
        
            # Ajout de Title dans le tableau
            $userObjet.Title = $GetUserObj.Title
 
            $userObjet.Phone = $GetUserObj.Phone
            $userObjet.MobilePhone = $GetUserObj.MobilePhone
 
            if ($GetUserObj.manager -ne $null)
            {
                $Manager = $GetUserObj.manager
                $objManager= $AllGetUserObj | where { $_.identity -like $Manager}
                $userObjet.manager = $ObjManager.LastName + " " + $ObjManager.FirstName

                $userObjet.manageremail = $ObjManager.WindowsEmailAddress
              
                
                #$userObjet.manager = ($AllGetUserObj | Where {$_.name -eq $_.manager}).DisplayName
                #$userObjet.manageremail = ($AllGetUserObj | Where {$_.name -eq $_.manager}).WindowsEmailAddress
 
                #Celui ok original
                #$userObjet.manager = $GetUserObj.manager
                #$userObjet.manageremail = ($AllGetUserObj | Where {$_.Name -eq $_.manager}).WindowsEmailAddress
                #
            }
            $userObjet.Identity = $GetUserObj.Identity
            $userObjet.Office = $GetUserObj.Office
            $userObjet.Department = $GetUserObj.Department
            $userObjet.CountryOrRegion = $GetUserObj.CountryOrRegion
            $userObjet.AccountDisabled = $GetUserObj.AccountDisabled 
#Gestion des images             
            $upn=$upn.toupper()
            if ($AllMailbox.UserprincipalName.toUpper() -ccontains $Upn)
           {
#                write-host "$(Get-Date) ### >> Recuperation compte O365....Traitement " -ForegroundColor Yellow        
                $PictPAth = $ImagesPath + $(($AllMailbox| where {$_.userprincipalName -eq $Upn}).WindowsEmailAddress -replace("@","-"))+".JPG"
                    $userObjet.ImageName = $($GetUserObj.WindowsEmailAddress -replace("@","-"))+".JPG"
                    # Backup File
                    if (Test-Path  $PictPAth) 
                    {               
                        if ($Refresh) 
                        {   
                            $Pict.PictureData |Set-Content $PictPath -Encoding byte 
                        }
                    }
                    Else 
                    { 
#                write-host "$(Get-Date) ### >> Recuperation Photo O365 et copie locale....Traitement " -ForegroundColor Yellow
                        $Pict = $null
                        $Pict =  Get-UserPhoto $Upn               
                        $Pict.PictureData |Set-Content $PictPath -Encoding byte 
                    }
                
       
            }
 
#fin de Gestion des images            
            
            
            
            
                   
        }
    Else
        {
            Write-Host $UPN " non trouve" -ForegroundColor DarkRed
        }  
 
   
        #$userObjet
        $AllUserObject+= $userObjet
 
   
 
}
write-host ""
write-host $(get-date)-ForegroundColor Yellow
# Disconnect 
write-host "$(Get-Date) ### >> Disconnect O365 - Exchange Session "  -ForegroundColor Yellow
Remove-PSSession -Session $SessionExchO365
 
 
write-host "$(Get-Date) ### >> Export User to CSV  "  -ForegroundColor Yellow
 
$AllUserObject | export-csv $CSVFile  -Delimiter ";" -Encoding UTF8 -NoTypeInformation
 
 
 
 
 
# Write HTML
 
#Region -- Build HTML
 
 
write-host "$(Get-Date) ### >> Customize HTML  "  -ForegroundColor Yellow 
 
 
#Custom heading?
If (Test-Path $HTMLPath\Heading.HTML)
{   write-host "$(Get-Date) ### >> Custom heading detected $HTMLPath\Heading.HTML.  Adding to Employee Directory"  -ForegroundColor Yellow
    $HeadingHTML = Get-Content $HTMLPath\Heading.HTML
}
Else
{   write-host "$(Get-Date) ### >> Custom $HeadingPath\Heading.HTML not found, will not be included in Employee Directory"  -ForegroundColor Yellow
}
#Custom CSS?
If (Test-Path $HTMLPath\CSS.HTML)
{   write-host "$(Get-Date) ### >> Custom CSS detected $HTMLPath\CSS.HTML overriding default CSS" -ForegroundColor Yellow
    $CSSHTML = Get-Content $HTMLPath\CSS.HTML
}
Else
{   write-host "$(Get-Date) ### >> No Custom CSS detected, using default" -ForegroundColor Yellow
    #Define Default CSS
    $CSSHTML = @"
 
table.Uncolored {
background:#ffffff; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#000000; 
 font:normal 12px Arial; 
 width:100%;
}
table.back {
background:#000000; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#FFFFFF; 
 font:normal 12px Arial; 
 width:100%;
}
 table.back  tr:hover {
background:#000000; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#FFFFFF; 
 font:normal 12px Arial; 
 width:100%;
}
 td.white {
      color:#FFFFFF; 
 }
table.Uncolored tr:hover { background:#FFFFFF; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#000000; 
 font:normal 12px Arial; 
 width:100%;
}  
     
form { 
  margin: 0; 
} 
table.TableStandard {
background:#e3e3e3; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#000000; 
 font:normal 12px Arial; 
 width:100%;
} 
 
td, th { 
 padding:.4em; 
} 
tr { border:1px dotted gray; 
} 
thead th, tfoot th { background:#000000; 
 color:#FFFFFF; 
 padding:3px 10px 3px 10px; 
 text-align:left; 
 text-transform:uppercase; 
} 
tbody th, tbody td { text-align:left; 
 vertical-align:top; 
} 
tbody tr:hover { background-color: #ffffff; 
 border:1px solid #123123; 
 color:#000000; 
}  
.styleHidePicture {
position:absolute;
visibility:hidden;
}
.styleShowPicture {
position:absolute;
visibility:visible;
border:solid 7px Black;
padding:1px;
}
/*mod*/
#lbBottomContainer {
overflow: visible;
}
/*/mod*/</style>
"@
}
$JSHTML = @"
<BODY onload="filter(document.getElementsByName('filt')[0], 'TableMain', '1');document.getElementsByName('filt')[0].focus();">
<script type='text/javascript'>
function filter (phrase, _id){
   var words = phrase.value.toLowerCase().split(" ");
   var table = document.getElementById(_id);
   var ele;
   var dark ;
   dark=0
   for (var r = 1; r < table.rows.length; r++){
         ele = table.rows[r].innerHTML.replace(/<[^>]+>/g,"");
           var displayStyle = 'none';
           for (var i = 0; i < words.length; i++) {
             if (ele.toLowerCase().indexOf(words[i])>=0)
               displayStyle = '';
             else {
               displayStyle = 'none';
               break;
             }
           }
         if (displayStyle==""){
           if (dark++){dark=0;table.rows[r].style.background= "#F7BFD9"}
           else {table.rows[r].style.background= "#64C2C7"}
         }
         table.rows[r].style.display = displayStyle;
         
   }
}
function filtbutton(phrase){
    var tableMain = document.getElementById('TableMain');
                   for(i=1;i<tableMain.rows.length;i++){
        ele = tableMain.rows[i].innerHTML.replace(/<[^>]+>/g,"");
        tableMain.rows[i].style.display = '';}
    for(i=1;i<tableMain.rows.length;i++){
                       ele = tableMain.rows[i].cells[4].innerHTML.replace(/<[^>]+>/g,"");
        if (ele != phrase && phrase != '') {
            tableMain.rows[i].style.display = 'none';
        } else {
        }
    }
}
<!--
function ShowPicture(id) {
                var currentDiv = document.getElementById(id);
                currentDiv.className='styleShowPicture'
}
function HidePicture(id) {
                var currentDiv = document.getElementById(id);
                currentDiv.className='styleHidePicture'
}
  //-->
</script>
 
 
 
 
"@ #End JSHTML
$HeaderHTML = @"
<HTML>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Expires" content="0" />
<TITLE>$Title</TITLE>
<script type="text/javascript" src="slimbox/js/mootools.js"></script>
<script type="text/javascript" src="slimbox/js/slimbox.js"></script>
<link rel="stylesheet" href="slimbox/css/slimbox.css" type="text/css" media="screen" />
<style type='text/css'>
 
"@  #End HeaderHTML
$EndHeaderHTML = @"
</style>
   
</HEAD>
"@ #End EndHeaderHTML
$SearchHTML = @"
 
 <Table class='back' bgcolor="000000">
      <tr>
             <td class='white'><marquee> $messageTxt </marquee> 
      </td>
      </tr>
</table>
      
<br>
<img src="Logo\logo.png" style="position:absolute">
<div style="text-align:center; height: 64px;"><font face="Tahoma" size="12">Annuaire</font></div>
 
<Table class='Uncolored'>
<tr>
      <TD><div align='left'><FORM action='javascript:void(0)'>
             <FONT face='Arial' size=2><b>Recherche: </b></FONT><input name='filt' onkeyup="filter(this, 'TableMain', '1')" type='text'></FORM></div>
             <br></TD>
      <!--mod-->
      <TD><div align='right'>
             <a href="infos.jpg" rel="lightbox"><b title="Informations CNIL">Informations CNIL</b></a></div>
      </div></TD>
      <!--/mod-->
</tr>
</Table>
 
 
 
"@ #End SearchHTML
$FooterHTML = @"
</TBODY>
</TABLE>
</BODY>
</HTML>
"@ #End FooterHTML
 
 
$GridHTMLA=@"
 
 
 
<TABLE id='TableMain' class='TableStandard'>
    <THEAD>
        
        <TR scope=""col"">
              <th>Nom</th>
              <th>Prenom</th>
              <th>Service</th>
              <th>Fonction</th>
              <th>Manager</th>
              <th>Emplacement</th>
              <th>Telephone</th>
              <th>Mobile</th>
              <th>Email</th>
        </TR>
    </THEAD>
 
    <TBODY> 
 
 
 
 
"@
 
$AllUserObject | sort LastName  | where {($_.Department -notlike "exclude" )} | where {($_.firstname -ne "" -and $_.lastname -ne "" ) }| % {
 
$GridHTML+=@"
    <tr>
        <td> 
"@
 
#title='" + $_.lastname + " "+ $_.Firstname+"'
if ($_.ImageName -ne $null ) { 
    $GridHTML+= "<a href='images/"+$_.ImageName+"' rel='lightbox' >"+[STRING]$_.LastName +"</a>"
 
}Else {
    $GridHTML+= [STRING]$_.LastName
 
}
 
 
$GridHTML+=@" 
        </td>
        <td> 
"@
$GridHTML+= $_.FirstName
$GridHTML+=@"
        </td>
        <td>  
"@
$GridHTML+= $_.Department
$GridHTML+=@" 
        </td>
        <td>  
"@
 
$GridHTML+= $_.Title
$GridHTML+=@"
        </td>
        <td>  
"@
 
$GridHTML+= $_.manager
$GridHTML+=@" 
        </td>
        <td>  
"@
$GridHTML+= $_.Office
$GridHTML+=@" 
        </td>
        <td>  
"@
$GridHTML+= $_.Phone
$GridHTML+=@" 
        </td>
        <td>  
"@
$GridHTML+= $_.MobilePhone
$GridHTML+=@" 
        </td>
        <td>  
"@
$GridHTML+= $_.email
$GridHTML+=@" 
</td>
"@
$GridHTML+=@"
 
    </tr>
 
"@    
 
 
}
 
$GridHTML =$GridHTMLA+$GridHTML
 
#EndRegion -- Build HTML
 
 
 
 
 
 
#Put it together and save
$HeadHTML = $HeaderHTML + $CSSHTML + $EndHeaderHTML + $JSHTML
$FullHTML = $HeadHTML + $HeadingHTML + $ButtonHTML  + $SearchHTML + $GridHTML + $FooterHTML
Write-host "$(Get-Date) ### >> Saving HTML: $htmlFile"  -ForegroundColor Yellow
$FullHTML | Out-File $htmlFile -Encoding UTF8
#& $OutputPath\index2.html                             #Un-remark if you wish to have the page displayed automatically in your browser
Write-host "$(Get-Date) ### >> Script completed"  -ForegroundColor GReen
 
 
 
$ErrorActionPreference = $ErrorActionPreferenceold
 
