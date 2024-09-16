  [CmdLetBinding()]
    Param (
      [Parameter ( Position = 0, Mandatory )][String]$SelectedArea,
      [Parameter ( Position = 1, Mandatory = $False )][ValidateSet( '0','1','2','5','10','15','30','50')][String]$RadiusInKM = 0,
      [Parameter ( Mandatory = $False )][ValidateSet('house','apartment','parking','land','storage_space','storage','berth','substructure','pitch')][String[]]$ObjectType = @('house','apartment'),
      [Parameter ( Mandatory = $False )][ValidateSet( '0','50.000','75.000','100.000','125.000','150.000','175.000','200.000','225.000','250.000','275.000','300.000','325.000','350.000','375.000','400.000','450.000','500.000','550.000','550.000','600.000','650.000','700.000','750.000','800.000','900.000','1.000.000','1.250.000','1.500.000','2.000.000','2.500.000','3.000.000','3.500.000','4.000.000','4.500.000','5.000.000' )][AllowEmptyString()][String]$PriceLow,
      [Parameter ( Mandatory = $False )][ValidateSet( '50.000','75.000','100.000','125.000','150.000','175.000','200.000','225.000','250.000','275.000','300.000','325.000','350.000','375.000','400.000','450.000','500.000','550.000','550.000','600.000','650.000','700.000','750.000','800.000','900.000','1.000.000','1.250.000','1.500.000','2.000.000','2.500.000','3.000.000','3.500.000','4.000.000','4.500.000','5.000.000' )][AllowEmptyString()][String]$PriceHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('garden','terrace','balcony')][String[]]$ExteriorSpaceType,
      [Parameter ( Mandatory = $False )][ValidateSet('north','west','south','east')][String[]]$ExteriorSpaceGardenOrientation,
      [Parameter ( Mandatory = $False )][ValidateSet('0','25','50','100','250','500')][AllowEmptyString()][String]$ExteriorSpaceGardenSizeLow,
      [Parameter ( Mandatory = $False )][ValidateSet('25','50','100','250','500')][AllowEmptyString()][String]$ExteriorSpaceGardenSizeHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('resale','newly_built')][String[]]$ConstructionType = @('resale','newly_built'),
      [Parameter ( Mandatory = $False )][ValidateSet('residential','recreational')][String[]]$Zoning = 'residential',
      [Parameter ( Mandatory = $False )][ValidateSet('before_1906','from_1906_to_1930','from_1931_to_1944','from_1945_to_1959','from_1960_to_1970','from_1971_to_1980','from_1981_to_1990','from_1991_to_2000','from_2001_to_2010','from_2011_to_2020','after_2020')][String[]]$ConstructionPeriod,
      [Parameter ( Mandatory = $False )][ValidateSet('in_residential_district','unobstructed_view','on_quiet_road','in_center','by_water','sheltered_position','on_navigable_waterway','open_position','rural','outside_built_up_area','on_busy_road','at_edge_of_woods','in_business_park','in_green_area','in_recreation_park','overlooking_park','sea_view')][String[]]$Surrounding,
      [Parameter ( Mandatory = $False )][ValidateSet('all_garages','lean_to','lock_up','garage_and_carport','built_in','underground','basement','detached','carport','parking_space','garage_possible')][String[]]$GarageType,
      [Parameter ( Mandatory = $False )][ValidateSet('0','1','2','3','4','5')][AllowEmptyString()][String]$GarageCapacityLow,
      [Parameter ( Mandatory = $False )][ValidateSet('1','2','3','4','5')][AllowEmptyString()][String]$GarageCapacityHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('lift','ground_floor','adapted_home','single_storey','accessible_for_the_elderly','accessible_for_the_disabled')][String[]]$Accessibility,
      [Parameter ( Mandatory = $False )][ValidateSet('bathtub','double_occupancy','renewable_energy','central_heating_boiler','fireplace','fixer_upper','swimming_pool')][String[]]$Amenities,
      [Parameter ( Mandatory = $False )][ValidateSet('single','group')][String[]]$Type,
      [Parameter ( Mandatory = $False )][ValidateSet('A+++++','A++++','A+++','A++','A+','A','B','C','D','E','F','G','H')][String[]]$EnergyLabel,
      [Parameter ( Mandatory = $False )][ValidateSet('0','1','2','3','4','5')][AllowEmptyString()][String]$BedroomsLow,
      [Parameter ( Mandatory = $False )][ValidateSet('1','2','3','4','5')][AllowEmptyString()][String]$BedroomsHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('0','1','2','3','4','5')][AllowEmptyString()][String]$RoomsLow,
      [Parameter ( Mandatory = $False )][ValidateSet('1','2','3','4','5')][AllowEmptyString()][String]$RoomsHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('0','250','500','750','1000','1500','2500','5000')][AllowEmptyString()][String]$PlotAreaLow,
      [Parameter ( Mandatory = $False )][ValidateSet('250','500','750','1000','1500','2500','5000')][AllowEmptyString()][String]$PlotAreaHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('0','50','75','100','125','150','175','200','250')][AllowEmptyString()][String]$FloorAreaLow,
      [Parameter ( Mandatory = $False )][ValidateSet('50','75','100','125','150','175','200','250')][AllowEmptyString()][String]$FloorAreaHigh,
      [Parameter ( Mandatory = $False )][ValidateSet('available','negotiations','unavailable')][String[]]$Availability = @('available','negotiations','unavailable'),
      [Parameter ( Mandatory = $False )][ValidateSet('1','3','5','10','30')][AllowEmptyString()][String]$PublicationDays,
      [Parameter ( Mandatory = $False )][ValidateSet('date_down','date_up','price_down','price_up','floor_area_down','floor_area_up','plot_area_down','plot_area_up','postal_code_up','relevancy')][String]$Sort = "date_down",
      [Parameter ( Mandatory = $False )][String]$FreeTextSearch,
      [Parameter ( Mandatory = $False )][Switch]$DownloadMedia,
      [Parameter ( Mandatory = $False )][Switch]$Refresh,
      [Parameter ( Mandatory = $False )][Int]$ResultCount
    )

  Begin {
    [System.Net.ServicePointManager]::SecurityProtocol = 3072 # Enable TLS 1.2 in PS Session
    $FundaSearchLink = 'https://www.funda.nl/zoeken/koop?selected_area=%5B"'+( ( $SelectedArea ).Replace(' ','-') )+','+$RadiusInKM+'km"%5D&object_type=%5B%22'+( $ObjectType -Join '%22,%22' )+'%22%5D'
    If ( $PriceLow -Or $PriceHigh ) { $FundaSearchLink = $FundaSearchLink + '&price=%22'+(( $PriceLow ).Replace('.',''))+'-'+(( $PriceHigh ).Replace('.',''))+'%22' }
    If ( $ExteriorSpaceType ) { $FundaSearchLink = $FundaSearchLink + '&exterior_space_type=%5B%22'+( $ExteriorSpaceType -Join '%22,%22' )+'%22%5D' }
    If ( $ExteriorSpaceGardenOrientation ) { $FundaSearchLink = $FundaSearchLink + '&exterior_space_garden_orientation=%5B%22'+( $ExteriorSpaceGardenOrientation -Join '%22,%22' )+'%22%5D' }
    If ( $ExteriorSpaceGardenSizeLow -Or $ExteriorSpaceGardenSizeHigh ) { $FundaSearchLink = $FundaSearchLink + "&exterior_space_garden_size=%22$ExteriorSpaceGardenSizeLow-$ExteriorSpaceGardenSizeHigh%22" }
    If ( $ConstructionType ) { $FundaSearchLink = $FundaSearchLink + '&construction_type=%5B%22'+( $ConstructionType -Join '%22,%22' )+'%22%5D' }
    If ( $Zoning ) { $FundaSearchLink = $FundaSearchLink + '&zoning=%5B%22'+( $Zoning -Join '%22,%22' )+'%22%5D' }
    If ( $ConstructionPeriod ) { $FundaSearchLink = $FundaSearchLink + '&construction_period=%5B%22'+( $ConstructionPeriod -Join '%22,%22' )+'%22%5D' }
    If ( $Surrounding ) { $FundaSearchLink = $FundaSearchLink + '&surrounding=%5B%22'+( $Surrounding -Join '%22,%22' )+'%22%5D' }
    If ( $GarageType ) { $FundaSearchLink = $FundaSearchLink + '&garage_type=%5B%22'+( $GarageType -Join '%22,%22' )+'%22%5D' }
    If ( $GarageCapacityLow -Or $GarageCapacityHigh ) { $FundaSearchLink = $FundaSearchLink + "&garage_capacity=%22$GarageCapacityLow-$GarageCapacityHigh%22" }
    If ( $Accessibility ) { $FundaSearchLink = $FundaSearchLink + '&accessibility=%5B%22'+( $Accessibility -Join '%22,%22' )+'%22%5D' }
    If ( $Amenities ) { $FundaSearchLink = $FundaSearchLink + '&amenities=%5B%22'+( $Amenities -Join '%22,%22' )+'%22%5D' }
    If ( $Type ) { $FundaSearchLink = $FundaSearchLink + '&type=%5B%22'+( $Type -Join '%22,%22' )+'%22%5D' }
    If ( $EnergyLabel ) { $FundaSearchLink = $FundaSearchLink + '&energy_label=%5B%22'+( ( $EnergyLabel -Replace "\+",'%2B' ) -Join '%22,%22' )+'%22%5D' }
    If ( $FreeTextSearch ) { $FundaSearchLink = $FundaSearchLink + "&free_text_search=%5B%22$FreeTextSearch%22%5D" }
    If ( $BedroomsLow -Or $BedroomsHigh ) { $FundaSearchLink = $FundaSearchLink + "&bedrooms=%22$BedroomsLow-$BedroomsHigh%22" }
    If ( $RoomsLow -Or $RoomsHigh ) { $FundaSearchLink = $FundaSearchLink + "&rooms=%22$RoomsLow-$RoomsHigh%22" }
    If ( $PlotAreaLow -Or $PlotAreaHigh ) { $FundaSearchLink = $FundaSearchLink + "&plot_area=%22$PlotAreaLow-$PlotAreaHigh%22" } # Perceeloppervlakte
    If ( $FloorAreaLow -Or $FloorAreaHigh ) { $FundaSearchLink = $FundaSearchLink + "&floor_area=%22$FloorAreaLow-$FloorAreaHigh%22" } # Woonoppervlakte
    If ( $Availability ) { $FundaSearchLink = $FundaSearchLink + '&availability=%5B%22'+( $Availability -Join '%22,%22' )+'%22%5D' }
    If ( $PublicationDays ) { $FundaSearchLink = $FundaSearchLink + "&publication_date=%22$PublicationDays%22" }
    If ( $Sort ) { $FundaSearchLink = $FundaSearchLink + "&sort=$Sort" }
    $OriginalWindowTitle = $Host.UI.RawUI.WindowTitle
    $NewWindowTitle = "Funda Scraper by D. Janmaat"
    $Host.UI.RawUI.WindowTitle = $NewWindowTitle
    $ProgressPreference = 'SilentlyContinue'
      @( 'ImportExcel' ).ForEach{
        If ( !( Get-Module $PSItem -ListAvailable ) ) {
          Install-Module $PSItem -Scope CurrentUser -Force -ErrorAction SilentlyContinue
          Import-Module $PSItem -ErrorAction SilentlyContinue
        }
      }
    $ScriptPath = Split-Path -Par $MyInvocation.MyCommand.Definition
    $ScriptBaseName = ( Get-Item $PSCommandPath ).Basename
    $Root = $PSScriptRoot
    $ExportResults = "$Root\Results_$ScriptBaseName.xlsx"
      If ( $Refresh ) {
        Remove-Item $ExportResults -Force -ErrorAction SilentlyContinue
        Remove-Item "$Root\MediaFolder" -Recurse -Force -ErrorAction SilentlyContinue
      }

    Function CatchString {
      [CmdLetBinding()]
        Param (
          [Parameter ( Position = 0, Mandatory )]$FirstString,
          [Parameter ( Position = 1, Mandatory )]$SecondString,
          [Parameter ( Position = 2, Mandatory )]$Content
        )
      Begin {
        $Content = $Content.Split([Environment]::NewLine)
        $Pattern = "$FirstString(.*?)$SecondString"
      }
      Process {
        $Result = [regex]::Match($Content,$Pattern).Groups[1].Value
      }
      End {
        Return $Result
        $FirstString = ""; $SecondString = ""; $Content = ""; $Result = ""
      }
    }

    Function RequestPage {
      [CmdLetBinding()]
        Param (
          [Parameter ( Position = 0, Mandatory )][String]$Uri
        )
      Begin{
        $Headers = @{
          "user-agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36 Edg/128.0.0.0";
          "authority" = "www.funda.nl";
          "accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7";
          "accept-encoding" = "gzip, deflate, br, zstd";
          "accept-language" = "en-US,en;q=0.9,nl;q=0.8";
          "cache-control" = "max-age=0";
        }
      }
      Process{
        Try {
          $Result = Invoke-WebRequest $Uri -Headers $Headers -UseBasicParsing -ErrorAction Stop
            If ( $Result.Content -Match 'We hebben last van een storing' ) {
              $Result = $False
            }
        } Catch {
          $Result = $False
        }
      }
      End{
        Return $Result
      }
    }

    Function TotalPages {
      [CmdLetBinding()]
        Param (
          [Parameter ( Position = 0, Mandatory )][String]$Uri
        )
      Begin {}
      Process {
        $Content = ( RequestPage $Uri ).Content
        $FirstString = '<a tabindex="0" data-v-b8a43de0>'
        $SecondString = '</a></li>'
        $Pattern = "$FirstString(.*?)$SecondString"
        $Links = ( $Content | Select-String -AllMatches $Pattern ).Matches.Value
          ForEach ( $Link in $Links | Where { $PSItem -NotMatch '>1</a>' -And $PSItem -Match '<a tabindex="0"' -And $PSItem -NotMatch 'http' } ) {
            $CatchString = CatchString '>' '</a>' ( [regex]::Replace($Link,"[^a-zA-Z0-9=<>/\s]","") )
              If ( $CatchString ) {
                $Result = $CatchString
                Break
              }
          }
      }
      End {
        Return [Int]$Result
      }
    }

    If ( !( RequestPage $FundaSearchLink ) ) {
      Write-Host "Error Funda : We hebben last van een storing`r`nCheck if selected area '$SelectedArea' is available on website https://funda.nl"
      Break
    }
    Clear-Host
    $TotalPageCount = TotalPages $FundaSearchLink
    If ( $TotalPageCount -Eq 0 ) { $TotalPageCount = 1 }
    Write-Host "Gather total houses..."
    $TotalLinkCount = ( CatchString 'Op funda vind je momenteel ' ' huizen te ' ( ( RequestPage $FundaSearchLink ).Content ) ).Replace( '.' , '' )
      If ( !$TotalLinkCount ) {
        Write-Host "Found no houses for selected area '$SelectedArea'"
        Break
      }
    Clear-Host
      If ( $ResultCount ) {
          If ( $ResultCount -Gt $TotalLinkCount ) {
            $ResultCount = $TotalLinkCount
          }
        $TotalLinkCount = $ResultCount  
      }
    Write-Host "Scraping $TotalLinkCount house(s) :`r`n"
    $ItemCounter = 0
  }

  Process {
    $TempExportResults = $Env:TEMP+'\'+$ScriptBaseName+'_temporary.xlsx'
      If ( Test-Path $ExportResults ) {
        $Sites = Import-Excel $ExportResults -WorksheetName "$ScriptBaseName"
          Try {
            $Sites | Export-Excel $TempExportResults -WorksheetName "$ScriptBaseName" -ErrorAction Stop
            Remove-Item $ExportResults -Force -ErrorAction SilentlyContinue
            Import-Excel $TempExportResults -WorksheetName "$ScriptBaseName" | Export-Excel $ExportResults -WorksheetName "$ScriptBaseName"
            Remove-Item $TempExportResults -Force -ErrorAction SilentlyContinue
            $BrokerInfo = Import-Excel $ExportResults -WorksheetName "$ScriptBaseName" | Select MakelaarId,Makelaar,MakelaarWebsite
          } Catch {
            Write-Host "Problem export existing results to temporary file $TempExportResults"
            Break
          }
      } Else {
        $Sites = @()
      }
    For ( $I = 1; $I -Le $TotalPageCount; $I++ ) {
      $Uri = "$FundaSearchLink&search_result=$I"
      $Invoke = RequestPage $Uri
      $Links = $Invoke.Links.href | Where { $PSItem -Match '/detail/' } | Select -Unique
        ForEach ( $Page in $Links ) {
          $ItemCounter++
            If ( $Sites.Page -NotContains $Page ) {
              $Content = ( RequestPage $Page ).Content
              $StraatHuisnummer = ( ( CatchString 'isinternational="' '</span><span' $Content ) -Split '>' )[-1]
              $Plaats = CatchString '"addressLocality":"' '"' $Content
              $StartString = 'ext-4xl">' + $StraatHuisnummer + '</span><span class="text-neutral-40">'
              $EndString = ' ' + $Plaats
              $Postcode = CatchString $StartString $EndString $Content
                If ( $Postcode.Length -Ge 10 ) {
                  $Postcode = ""
                }
                If ( $Content -Match '"addressRegion":"' ) {
                  $Provincie = CatchString '"addressRegion":"' '"' $Content
                } Else {
                  $Provincie = ""
                }
              $Plaats = ( $Plaats ).Replace( ' / ' , '-' ).Replace( '/ ' , '-' ).Replace( ' - ' , '-' )
              $BuurtIdentifier = [cultureinfo]::GetCultureInfo("nl-NL").TextInfo.ToTitleCase( ( ( CatchString 'neighborhoodidentifier="' '"' $Content ) -Split '/' )[-1] )
                If ( !( $BuurtIdentifier ) ) {
                  $PlaatsBuurt = $Plaats
                } Else {
                  $PlaatsBuurt = $Plaats+' - '+$BuurtIdentifier
                }
              $Land = CatchString 'country="' '" isinternational=' $Content
              $VolledigAdres = ""
              $StraatHuisnummer = ( $StraatHuisnummer ).Replace( '&#39;',"'" )
              If ( $StraatHuisnummer ) { $VolledigAdres = $VolledigAdres + $StraatHuisnummer }
              If ( $Postcode ) { $VolledigAdres = $VolledigAdres + ', ' + $Postcode }
              If ( $PlaatsBuurt ) { $VolledigAdres = $VolledigAdres + ', ' + $PlaatsBuurt }
              If ( $Provincie ) { $VolledigAdres = $VolledigAdres + ', ' + $Provincie }
              If ( $Land ) { $VolledigAdres = $VolledigAdres + ', ' + $Land }
              $Longitude = ( ( CatchString '"Longitude":' ',"' $Content ) -Split ',' )[-1]
              $Latitude = ( ( CatchString '"Longitude":' ',"' $Content ) -Split ',' )[-2]
              $GoogleMaps = 'https://www.google.nl/maps/place/' + ( CatchString 'https://www.google.nl/maps/place/' '",' $Content )
              $Beschrijving = CatchString '<meta name="description" content="' '">' $Content
              $Internationaal = CatchString 'isinternational="' '"' $Content; If ( $Internationaal -Eq 'false' ) { $Internationaal = "Nee" } Else { $Internationaal = "Ja" }
              $MakelaarId = ( ( CatchString '"listing_place"' '","/detail/koop/' $Content ).Split( '},' )[-1] ).Split(',"')[0]
                ( $BrokerInfo ).ForEach{
                  If ( $PSItem.MakelaarId -Eq $MakelaarId ) {
                    $Makelaar = $PSItem.Makelaar
                    $MakelaarWebsite = $PSItem.MakelaarWebsite
                  }
                }
                If ( !$Makelaar ) {
                  $MakelaarRequest = ( RequestPage "https://funda.nl/makelaar/$MakelaarId" )
                  $Makelaar = ( ( CatchString '<title>' '</title>' ( $MakelaarRequest ).Content ).Replace(' [funda]' , '' ).Replace( '&amp;' , '&' ).Replace( '&#x27;' , "`'" ).Replace( '&#x2F;' , '-' ) ).Trim()
                  $MakelaarWebSite = ( ( $MakelaarRequest ).Links.href | Where { $PSItem -NotMatch 'funda' -And $PSItem -NotMatch 'facebook' -And $PSItem -NotMatch 'instagram' -And $PSItem -Match 'http' -And $PSItem -NotMatch 'nvm.nl' -And $PSItem -NotMatch 'x.com' -And $PSItem -NotMatch 'linkedin.com' -And $PSItem -NotMatch 'twitter' -And $PSItem -NotMatch 'wa.me' -And $PSItem -NotMatch 'vastgoedpro.nl' -And $PSItem -NotMatch 'vbomakelaar.nl' } | Select -Unique ) -Join ','
                  #$MakelaarTelefoonnummer = CatchString 'tel:' '"' $MakelaarRequest
                  #$MakelaarEmailadres = ( ( CatchString '{"email":' '","' $MakelaarRequest ) -Split '"' )[-1]
                  #$MakelaarAdres = ( CatchString '<address class="not-italic">' '</address>' $MakelaarRequest ).Replace( ' <br> ' , ', ' )
                  Write-Host "New broker found : $Makelaar"
                  If ( Test-Path $ExportResults ) { $BrokerInfo = Import-Excel $ExportResults -WorksheetName "$ScriptBaseName" | Select MakelaarId,Makelaar,MakelaarWebsite }
                }
              Write-Host "$ItemCounter / $TotalLinkCount - Scraping : $Makelaar - $VolledigAdres"
              $AangebodenSinds = CatchString '"Aangeboden sinds","' '",{' $Content
              $LaatsteVraagprijs = ""
              $Vraagprijs = ( ( ( CatchString 'Vraagprijs</dt>' '</span>' $Content ) -Split '>' )[-1] ).Replace( "`€ " , '' ).Replace( '.' , '' )
                If ( !$Vraagprijs ) {
                  $LaatsteVraagprijs = ( CatchString '"Laatste vraagprijs","' '","' $Content ).Replace( "`€ " , '' ).Replace( '.' , '' )
                } Else { $LaatsteVraagprijs = "" }
              $KostenKoperString = ' kosten koper'
              $VrijOpNaamString = ' vrij op naam'
              $Koopvorm = ''
                If ( $Vraagprijs -And $Vraagprijs -Match $KostenKoperString ) {
                  $Vraagprijs = ( $VraagPrijs ).Replace( $KostenKoperString, '' )
                  $Koopvorm = 'Kosten Koper'
                }
                If ( $LaatsteVraagprijs -And $LaatsteVraagprijs -Match $KostenKoperString ) {
                  $LaatsteVraagprijs = ( $LaatsteVraagPrijs ).Replace( $KostenKoperString, '' )
                  $Koopvorm = 'Kosten Koper'
                }
                If ( $Vraagprijs -And $Vraagprijs -Match $VrijOpNaamString ) {
                  $Vraagprijs = ( $VraagPrijs ).Replace( $VrijOpNaamString, '' )
                  $Koopvorm = 'Vrij Op Naam'
                }
                If ( $LaatsteVraagprijs -And $LaatsteVraagprijs -Match $VrijOpNaamString ) {
                  $LaatsteVraagprijs = ( $LaatsteVraagPrijs ).Replace( $VrijOpNaamString, '' )
                  $Koopvorm = 'Vrij Op Naam'
                }
              $VraagprijsperM = ( ( ( CatchString 'Vraagprijs per m²</dt>' '</span>' $Content ) -Split '>' )[-1] ).Replace( "`€ " , '' ).Replace( '.' , '' )
              $VerkoopDatum = ( ( CatchString 'Verkoopdatum</dt' '</dd>' $Content ) -Split '>' )[-1]
              $Status = ( ( CatchString 'Status</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Aanvaarding = ( ( CatchString 'Aanvaarding</dt>' '</span>' $Content ) -Split '>' )[-1]
              $SoortWoning = ( ( CatchString 'Soort woonhuis</dt>' '</span>' $Content ) -Split '>' )[-1]
              $SoortBouw = ( ( CatchString 'Soort bouw</dt>' '</span>' $Content ) -Split '>' )[-1]
              $BouwJaar = ( ( CatchString 'Bouwjaar</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Toegankelijkheid = ( ( CatchString 'Toegankelijkheid</dt>' '</span>' $Content ) -Split '>' )[-1]
              $SoortDak = ( ( CatchString 'Soort dak</dt>' '</span>' $Content ) -Split '>' )[-1]
              $GebouwgebondenBuitenruimte = CatchString '"Gebouwgebonden buitenruimte","' ' m²"' $Content
              $ExterneBergruimte = CatchString '"Externe bergruimte","' ' m²"' $Content
              $Woonoppervlakte = ( ( CatchString '/kaart" class="mt' 'm²' $Content ) -Split '"md:font-bold">' )[-1]
              $Perceel = ( ( ( CatchString 'Perceel</dt>' '</span>' $Content ) -Split '>' )[-1] ).Replace( ' m²' , '' ).Replace( '.' , '' )
              $Inhoud = ( ( ( CatchString 'Inhoud</dt>' '</span>' $Content ) -Split '>' )[-1] ).Replace( ' m³' , '' ).Replace( '.' , '' )
              $AantalKamers = ( ( CatchString 'Aantal kamers</dt>' '</span>' $Content ) -Split '>' )[-1]
              $AantalBadkamers = ( ( CatchString 'Aantal badkamers</dt>' '</span>' $Content ) -Split '>' )[-1]
              $BadkamerVoorzieningen = ( ( CatchString 'Badkamervoorzieningen</dt>' '</span>' $Content ) -Split '>' )[-1]
              $AantalWoonlagen = ( ( CatchString 'Aantal woonlagen</dt>' '</span>' $Content ) -Split '>' )[-1]
              $EnergieLabel = ( ( CatchString 'Energielabel</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Isolatie = ( ( CatchString 'Isolatie</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Verwarming = ( ( CatchString 'Verwarming</dt>' '</span>' $Content ) -Split '>' )[-1]
              $WarmWater = ( ( CatchString 'Warm water</dt>' '</span>' $Content ) -Split '>' )[-1]
              $KadastraleGegevens = ( ( CatchString 'Kadastrale gegevens</h3>' ' <!----></dt>' $Content ) -Split '>' )[-1]
              $KadastraleGegevensLink = 'https://www.funda.nl' + ( ( CatchString $KadastraleGegevens '" class=' $Content ) -Split '<a href="' )[-1]
                If ( $KadastraleGegevensLink -Match 'noopener noreferrer' ) {
                  $KadastraleGegevensLink = "Geen gegevens"
                }
              $Oppervlakte = ( ( ( CatchString 'Oppervlakte</dt>' '</dd>' $Content ) -Split '>' )[-1] ).Replace( ' m²' , '' ).Replace( '.' , '' )
              $EigendomsSituatie = ( ( CatchString 'Eigendomssituatie</dt>' '</dd>' $Content ) -Split '>' )[-1]
              $Ligging = ( ( CatchString 'Ligging</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Tuin = ( ( CatchString 'Tuin</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Achtertuin = ( ( CatchString 'Achtertuin</dt>' '</span>' $Content ) -Split '>' )[-1]
                If ( $Achtertuin -Match 'm²' ) {
                  $Temp = $Achtertuin -Split ' m²'
                  $AchtertuinM = $Temp[0]
                } Else { $AchtertuinM = "" }
              $LiggingTuin = ( ( CatchString 'Ligging tuin</dt>' '</span>' $Content ) -Split '>' )[-1]
              $SchuurBerging = ( ( CatchString 'Schuur/berging</dt>' '</span>' $Content ) -Split '>' )[-1]
                If ( $SchuurBerging ) {
                  $Temp = ( ( CatchString 'Schuur/berging</dt>' 'class="mt-4 font-bold">' $Content ) -Split 'Voorzieningen</dt>' )[1]
                    If ( $Temp ) {
                      $SchuurBergingVoorzieningen = ( ( CatchString '<dd class=' '</span>' $Temp ) -Split '>')[-1]
                    } Else { $SchuurBergingVoorzieningen = "" }
                } Else { $SchuurBergingVoorzieningen = "" }
              $SoortGarage = ( ( CatchString 'Soort garage</dt>' '</span>' $Content ) -Split '>' )[-1]
              $GarageCapaciteit = ( ( ( CatchString 'Capaciteit</dt>' '</span>' $Content ) -Split '>' )[-1] ).Replace( '&#39;' , "`'" )
                If ( $GarageCapaciteit ) {
                  $GarageVoorzieningen = ( ( ( CatchString 'Capaciteit</dt>' '</span><!' $Content ) -Split '>' )[-1] ).Replace( '&#39;' , "`'" )
                } Else { $GarageVoorzieningen = "" }
              $SoortParkeerGelegenheid = ( ( CatchString 'Soort parkeergelegenheid</dt>' '</span>' $Content ) -Split '>' )[-1]
              $Omschrijving = ( ( ( CatchString '"Omschrijving","' '"' $Content ) -Split '>' )[-1] ).Replace( ' \n\n' , ' ' ).Replace( ' \n' , '. ' ).Replace( '\n' , ' ' )
              $DateTimeScraped = ( Get-Date ).ToString( "d MMMM yyyy HH:mm:ss",[CultureInfo]"nl-NL" )
              $Object = New-Object PSCustomObject
              $Object | Add-Member 'Page' $Page
              $Object | Add-Member 'VolledigAdres' $VolledigAdres
              $Object | Add-Member 'AangebodenSinds' $AangebodenSinds
              $Object | Add-Member 'Vraagprijs' $Vraagprijs
              $Object | Add-Member 'LaatsteVraagprijs' $LaatsteVraagprijs
              $Object | Add-Member 'Koopvorm' $Koopvorm
              $Object | Add-Member 'VraagprijsperM2' $VraagprijsperM
              $Object | Add-Member 'Status' $Status
              $Object | Add-Member 'Makelaar' $Makelaar
              $Object | Add-Member 'BouwJaar' $BouwJaar
              $Object | Add-Member 'GebouwgebondenBuitenruimteM2' $GebouwgebondenBuitenruimte
              $Object | Add-Member 'ExterneBergruimteM2' $ExterneBergruimte
              $Object | Add-Member 'WoonoppervlakteM2' $Woonoppervlakte
              $Object | Add-Member 'PerceelM2' $Perceel
              $Object | Add-Member 'InhoudM3' $Inhoud
              $Object | Add-Member 'Aanvaarding' $Aanvaarding
              $Object | Add-Member 'SoortWoning' $SoortWoning
              $Object | Add-Member 'SoortBouw' $SoortBouw
              $Object | Add-Member 'Toegankelijkheid' $Toegankelijkheid
              $Object | Add-Member 'SoortDak' $SoortDak
              $Object | Add-Member 'AantalKamers' $AantalKamers
              $Object | Add-Member 'AantalBadkamers' $AantalBadkamers
              $Object | Add-Member 'BadkamerVoorzieningen' $BadkamerVoorzieningen
              $Object | Add-Member 'AantalWoonlagen' $AantalWoonlagen
              $Object | Add-Member 'EnergieLabel' $EnergieLabel
              $Object | Add-Member 'Isolatie' $Isolatie
              $Object | Add-Member 'Verwarming' $Verwarming
              $Object | Add-Member 'WarmWater' $WarmWater
              $Object | Add-Member 'OppervlakteM2' $Oppervlakte
              $Object | Add-Member 'EigendomsSituatie' $EigendomsSituatie
              $Object | Add-Member 'Ligging' $Ligging
              $Object | Add-Member 'Tuin' $Tuin
              $Object | Add-Member 'Achtertuin' $Achtertuin
              $Object | Add-Member 'AchtertuinM' $AchtertuinM
              $Object | Add-Member 'LiggingTuin' $LiggingTuin
              $Object | Add-Member 'SchuurBerging' $SchuurBerging
              $Object | Add-Member 'SchuurBergingVoorzieningen' $SchuurBergingVoorzieningen
              $Object | Add-Member 'SoortGarage' $SoortGarage
              $Object | Add-Member 'GarageCapaciteit' $GarageCapaciteit
              $Object | Add-Member 'GarageVoorzieningen' $GarageVoorzieningen
              $Object | Add-Member 'SoortParkeerGelegenheid' $SoortParkeerGelegenheid
              $Object | Add-Member 'StraatHuisnummer' $StraatHuisnummer
              $Object | Add-Member 'Postcode' $Postcode
              $Object | Add-Member 'Plaats' $Plaats
              $Object | Add-Member 'Provincie' $Provincie
              $Object | Add-Member 'Land' $Land
              $Object | Add-Member 'BuurtIdentifier' $BuurtIdentifier
              $Object | Add-Member 'Internationaal' $Internationaal
              $Object | Add-Member 'Longitude' $Longitude
              $Object | Add-Member 'Latitude' $Latitude
              $Object | Add-Member 'GoogleMaps' $GoogleMaps
              $Object | Add-Member 'KadastraleGegevens' $KadastraleGegevens
              $Object | Add-Member 'KadastraleGegevensLink' $KadastraleGegevensLink
              $Object | Add-Member 'MakelaarId' $MakelaarId
              $Object | Add-Member 'MakelaarWebsite' $MakelaarWebsite
              #$Object | Add-Member 'MakelaarAdres' $MakelaarAdres
              #$Object | Add-Member 'MakelaarTelefoonnummer' $MakelaarTelefoonnummer
              #$Object | Add-Member 'MakelaarEmailadres' $MakelaarEmailadres
              $Object | Add-Member 'Beschrijving' $Beschrijving
              $Object | Add-Member 'Omschrijving' $Omschrijving
              $Object | Add-Member 'DateTimeScraped' $DateTimeScraped
              $Object | Add-Member 'Print' ( $Page+'/print/' )
              $TryCount = 0
              $TryCountEnd = 20
              $RetryInterval = 1500
                Do {
                  $Saved = $False
                  Try {
                    $Object | Export-Excel $ExportResults -WorksheetName "$ScriptBaseName" -Append -ErrorAction Stop
                    $Saved = $True
                  } Catch {
                    $TryCount++
                    Write-Host "Problem saving, retry $TryCount / $TryCountEnd"
                    Start-Sleep -M $RetryInterval
                  }
                } Until ( $Saved -Or $Count -Ge $TryCountEnd )
                If ( $DownloadMedia ) {
                  $MediaBaseFolder = "$Root\MediaFolder"; New-Item $MediaBaseFolder -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                  $MediaPageFolder = "$MediaBaseFolder\$VolledigAdres"; New-Item $MediaPageFolder -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
                  $Media = @()
                  $FirstString = '"contentUrl":"'
                  $SecondString = '"}'
                  $Pattern = "$FirstString(.*?)$SecondString"
                  $ImageLinks = ( $Content | Select-String -AllMatches $Pattern ).Matches.Value
                    ForEach ( $ImageLink in $ImageLinks ) {
                      $ImageLink = ( $ImageLink ).Replace( '"contentUrl":"' , '' ).Replace( '"}' , '' )
                      $FileName = Split-Path $ImageLink -Leaf
                      $Object = New-Object PSCustomObject
                      $Object | Add-Member 'Name' $FileName
                      $Object | Add-Member 'Link' $ImageLink
                      $Media += $Object
                    }
                  $MediaFiles = ( RequestPage $Page ).Links.href | Where {$PSItem -Match 'cloud.funda.nl'} | Select -Unique
                    ForEach ( $MediaFileLink in $MediaFiles ) {
                      $FileName = Split-Path $MediaFileLink -Leaf
                      $Object = New-Object PSCustomObject
                      $Object | Add-Member 'Name' $FileName
                      $Object | Add-Member 'Link' $MediaFileLink
                      $Media += $Object
                    }
                    If ( $Media ) {
                      $Media | ForEach -Parallel {
                        $MediaName = $PSItem.Name
                        $MediaLink = $PSItem.Link
                        $MediaFullName = $Using:MediaPageFolder+'\'+$MediaName
                        Write-Host "Scraping media : download media file $MediaName"
                        Invoke-WebRequest $MediaLink -UseBasicParsing -OutFile $MediaFullName -ErrorAction SilentlyContinue
                      } -AsJob | Out-Null
                    } Else {
                      Write-Host "Scraping media : no media found"
                    }
                }
            } Else {
              $AlreadyScraped = Import-Excel $ExportResults -WorksheetName "$ScriptBaseName" | Where { $PSItem.Page -Eq $Page } | Select VolledigAdres,Makelaar
              $AlreadyScrapedAddress = $AlreadyScraped.VolledigAdres
              $AlreadyScrapedBroker = $AlreadyScraped.Makelaar
              Write-Host "$ItemCounter / $TotalLinkCount - Already scraped : $AlreadyScrapedBroker - $AlreadyScrapedAddress"
            }
          $DateTimeScraped = ""; $Omschrijving = ""; $SoortParkeerGelegenheid = ""; $GarageVoorzieningen = ""; $GarageCapaciteit = ""; $SoortGarage = ""
          $SchuurBergingVoorzieningen = ""; $SchuurBerging = ""; $LiggingTuin = ""; $Tuin = ""; $Achtertuin = ""; $Ligging = ""; $EigendomsSituatie = ""
          $Oppervlakte = ""; $KadastraleGegevensLink = ""; $KadastraleGegevens = ""; $WarmWater = ""; $Verwarming = ""; $Isolatie = ""; $EnergieLabel = ""
          $AantalWoonlagen = ""; $BadkamerVoorzieningen = ""; $AantalBadkamers = ""; $SoortDak = ""; $Toegankelijkheid = ""; $SoortBouw = ""; $Item = ""
          $SoortWoning = ""; $Aanvaarding = ""; $MakelaarWebsite = ""; $MakelaarId = ""; $GoogleMaps = ""; $Beschrijving = ""; $Latitude = ""
          $Longitude = ""; $Internationaal = ""; $Inhoud = ""; $Perceel = ""; $Woonoppervlakte = ""; $ExterneBergruimte = ""; $Bouwjaar = ""; $AchtertuinM = ""
          $GebouwgebondenBuitenruimte = ""; $Status = ""; $VraagprijsperM = ""; $Koopvorm = ""; $LaatsteVraagprijs = ""; $Vraagprijs = ""
          $AangebodenSinds = ""; $Makelaar = ""; $BuurtIdentifier = ""; $Land = ""; $Provincie = ""; $Plaats = ""; $Postcode = ""; $StraatHuisnummer = ""
            If ( $ItemCounter -Eq $ResultCount ) {
              Break
            }
        }
      If ( $ItemCounter -Eq $ResultCount ) {
        Break
      }
    }
  }

  End {
    $ExcelPackage = Export-Excel $ExportResults -WorkSheetName "$ScriptBaseName" -TableStyle 'Medium2' -AutoSize -FreezePane 2,4 -BoldTopRow -PassThru
    $PivotNameVraagprijsPerM = "Vraagprijs per M2"
    Add-PivotTable -ExcelPackage $ExcelPackage -PivotRows Makelaar,VolledigAdres -PivotData @{ "VraagprijsperM2" = "Average" } -SourceWorksheet "$ScriptBaseName" -PivotTableName $PivotNameVraagprijsPerM
    @(4,5,7).ForEach{ Set-ExcelColumn -WorkSheet $ExcelPackage."$ScriptBaseName" -Column $PSItem -NumberFormat "[$€-nl-NL] #,##0" -HorizontalAlignment Left -VerticalAlignment Top }
    Set-ExcelRow -Worksheet $ExcelPackage."$ScriptBaseName" -Row 1 -TextRotation 60 -Height 75 -HorizontalAlignment Left -VerticalAlignment Top -WrapText
    Set-ExcelRow -WorkSheet $ExcelPackage."$ScriptBaseName" -Row 2 -Height 15
    For( $Column = 1; $Column -Le 60; $Column++ ) { Set-ExcelColumn -WorkSheet $ExcelPackage."$ScriptBaseName" -Column $Column -FontName Calibri -FontSize 11 -HorizontalAlignment Left -VerticalAlignment Top }
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range A:A -Width 10
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range C:C -Width 17
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range D:E -Width 12
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range F:F -Width 12
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range G:G -Width 10
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range G:G -Width 9
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range J:J -Width 7
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range K:O -Width 12
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range Y:Y -Width 15
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range AC:AC -Width 10
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range AD:AD -Width 14
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range AH:AH -Width 8
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range AM:AN -Width 9
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range AW:AX -Width 8
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range BB:BB -Width 7
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range BD:BD -Width 50
    Set-ExcelRange -Wo $ExcelPackage."$ScriptBaseName" -Range BE:BE -Width 135
    Close-ExcelPackage $ExcelPackage
    $ExcelPackage.Dispose()
    Write-Host "`r`nScraped $ItemCounter site(s) from Funda`r`n"
    $Host.UI.RawUI.WindowTitle = $OriginalWindowTitle
  }
