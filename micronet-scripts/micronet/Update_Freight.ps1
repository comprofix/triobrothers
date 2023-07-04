if ($env:Processor_Architecture -ne "x86")
{ write-warning 'Launching x86 PowerShell'
&"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NoProfile $myInvocation.Line -executionpolicy unrestricted
exit
}

$MYDSN = "Micronet"

function addfreight($freightno,$DEL_DELNO,$del_dbtno) {
  $query = "UPDATE Debtors_Delivery_Address_File SET DEL_FRGTNO = '$freightno' WHERE DEL_DELNO = '$DEL_DELNO' AND DEL_DBTNO = '$del_dbtno'"
  Write-Host "Updating $del_dbtno Address Number $DEL_DELNO to have $freightno Freight Code - Post Code: $DEL_POSTCODE. Suburb: $DEL_DELADR2"
  $cmd = new-object System.Data.Odbc.OdbcCommand($query,$conn)
  $cmd.ExecuteNonQuery() | Out-Null

}

$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;Option=67108864;"
$conn = New-Object System.Data.Odbc.OdbcConnection
$conn.ConnectionString = $dsn
$conn.Open()
clear-Host

$query = "SELECT DEL_DBTNO, DEL_DELNO, DEL_DELADR0, DEL_DELADR1, DEL_DELADR2, DEL_DELADR3, DEL_DELADR4, DEL_POSTCODE, DEL_FRGTNO, DBT_STATUS, DBT_CLASS FROM Debtors_Delivery_Address_File, Debtors_Master_File WHERE Debtors_Master_File.DBT_NO = Debtors_Delivery_Address_File.DEL_DBTNO AND DBT_STATUS = '0' AND DEL_FRGTNO IS NULL AND (DBT_CLASS LIKE 'OTHER%' OR DBT_CLASS LIKE 'PUMEN%' OR DBT_CLASS LIKE 'PAWN%' OR DBT_CLASS LIKE 'RC%' OR DBT_CLASS LIKE 'TC%')"
$cmd = (New-object System.Data.Odbc.OdbcCommand($query,$conn)).ExecuteReader()
$table = new-object "System.Data.DataTable"
$table.Load($cmd)
#$table | Out-Gridview -wait

foreach ($address in $table) {
  $del_dbtno = $address.DEL_DBTNO
  $DEL_DELNO = $address.DEL_DELNO
  $DEL_POSTCODE = $address.DEL_POSTCODE
  $DEL_FRGTNO = $address.DEL_FRGTNO
  $DEL_DELADR2 = $address.DEL_DELADR2

  addfreight FREE $DEL_DELNO $del_dbtno
}

$query = "SELECT DEL_DBTNO, DEL_DELNO, DEL_DELADR0, DEL_DELADR1, DEL_DELADR2, DEL_DELADR3, DEL_DELADR4, DEL_POSTCODE, DEL_FRGTNO, DBT_STATUS, DBT_CLASS FROM Debtors_Delivery_Address_File, Debtors_Master_File WHERE Debtors_Master_File.DBT_NO = Debtors_Delivery_Address_File.DEL_DBTNO AND DBT_STATUS = '0' AND DEL_FRGTNO IS NULL AND (DBT_CLASS NOT LIKE 'TC%' OR DBT_CLASS NOT LIKE 'PUM%' OR DBT_CLASS NOT LIKE 'RC%' OR DBT_CLASS NOT LIKE 'O/S%')"
$cmd = (New-object System.Data.Odbc.OdbcCommand($query,$conn)).ExecuteReader()
$table = new-object "System.Data.DataTable"
$table.Load($cmd)
#$table | Out-Gridview -wait

foreach ($address in $table) {

  $del_dbtno = $address.DEL_DBTNO
  $DEL_DELNO = $address.DEL_DELNO
  $DEL_POSTCODE = $address.DEL_POSTCODE
  $DEL_FRGTNO = $address.DEL_FRGTNO
  $DEL_DELADR2 = $address.DEL_DELADR2

  #SYNCBD
  If ($DEL_POSTCODE -In 2000..2009) {
    addfreight SYNCBD $DEL_DELNO $del_dbtno
  }

  #SYDMET
  If ($DEL_POSTCODE -In 2010..2103 -or
      $DEL_POSTCODE -In 2110..2155 -or
      $DEL_POSTCODE -In 2160..2173 -or
      $DEL_POSTCODE -In 2176..2232) {
    addfreight SYDMET $DEL_DELNO $del_dbtno
  }

  #SYDOUT
  If ($DEL_POSTCODE -In 2104..2109 -or
      $DEL_POSTCODE -In 2156..2159 -or
      $DEL_POSTCODE -In 2174..2175 -or
      $DEL_POSTCODE -In 2233..2234 -or
      $DEL_POSTCODE -In 2557..2570 -or
      $DEL_POSTCODE -In 2745..2770) {
        addfreight SYDOUT $DEL_DELNO $del_dbtno
  }

  #WOLLO
  If ($DEL_POSTCODE -In 2500..2506) {
    addfreight WOLLO $DEL_DELNO $del_dbtno
  }

  #NEWC
  If ($DEL_POSTCODE -In 2298..2300 -or
      $DEL_POSTCODE -In 2302..2308) {
    addfreight NEWC $DEL_DELNO $del_dbtno
  }

  #Canberra
  If ($DEL_POSTCODE -In 2600..2620 -or
      $DEL_POSTCODE -In 2900..2914) {
        addfreight Canberra $DEL_DELNO $del_dbtno
  }

  #N1
  If ($DEL_POSTCODE -In 2250..2297 -or
      $DEL_POSTCODE -eq 2301 -or
      $DEL_POSTCODE -In 2309..2327 -or
      $DEL_POSTCODE -In 2507..2530 -or
      $DEL_POSTCODE -In 2773..2786 ) {
        addfreight N1 $DEL_DELNO $del_dbtno
  }

  #N2
  If ($DEL_POSTCODE -In 2328..2340 -or
      $DEL_POSTCODE -In 2422..2453 -or
      $DEL_POSTCODE -In 2533..2535 -or
      $DEL_POSTCODE -In 2538..2541 -or
      $DEL_POSTCODE -In 2571..2580 -or
      $DEL_POSTCODE -In 2621..2625 -or
      $DEL_POSTCODE -In 2787..2820 -or
      $DEL_POSTCODE -In 2830..2831 -or
      $DEL_POSTCODE -In 2844..2871) {
        addfreight N2 $DEL_DELNO $del_dbtno
  }

  #N3
  If ($DEL_POSTCODE -In 2341..2421 -or
      $DEL_POSTCODE -In 2454..2490 -or
      $DEL_POSTCODE -In 2536..2537 -or
      $DEL_POSTCODE -In 2545..2551 -or
      $DEL_POSTCODE -In 2581..2594 -or
      $DEL_POSTCODE -In 2626..2710 -or
      $DEL_POSTCODE -In 2720..2739 -or
      $DEL_POSTCODE -In 2821..2829 -or
      $DEL_POSTCODE -In 2840..2843 -or
      $DEL_POSTCODE -In 2873..2874) {
        addfreight N3 $DEL_DELNO $del_dbtno
  }

  #N4
  If ($DEL_POSTCODE -In 2711..2717 -or
      $DEL_POSTCODE -In 2832..2839 -or
      $DEL_POSTCODE -In 2875..2880) {
        addfreight N4 $DEL_DELNO $del_dbtno
  }


  #ADEMET
  If ($DEL_POSTCODE -In 5000..5115 -or
      $DEL_POSTCODE -In 5121..5127 -or
      $DEL_POSTCODE -In 5158..5169 -or
      $DEL_POSTCODE -eq 5950) {
        addfreight ADEMET $DEL_DELNO $del_dbtno
  }

  #SA A
  If ($DEL_POSTCODE -In 5116..5118 -or
      $DEL_POSTCODE -In 5131..5157 -or
      $DEL_POSTCODE -In 5170..5174 -or
      $DEL_POSTCODE -In 5201..5202 -or
      $DEL_POSTCODE -In 5231..5232 -or
      $DEL_POSTCODE -In 5245..5251) {
        addfreight SAA $DEL_DELNO $del_dbtno
  }

  #SA B
  If ($DEL_POSTCODE -eq 5120 -or
      $DEL_POSTCODE -In 5203..5214 -or
      $DEL_POSTCODE -In 5233..5244 -or
      $DEL_POSTCODE -In 5252..5261 -or
      $DEL_POSTCODE -In 5350..5413 -or
      $DEL_POSTCODE -In 5460..5461 -or
      $DEL_POSTCODE -In 5501..5510 -or
      $DEL_POSTCODE -In 5550..5558 -or
      $DEL_POSTCODE -In 5570..5583) {
        addfreight SAB $DEL_DELNO $del_dbtno
  }

  #SA C
  If ($DEL_POSTCODE -In 5262..5346 -or
      $DEL_POSTCODE -In 5414..5433 -or
      $DEL_POSTCODE -In 5451..5455 -or
      $DEL_POSTCODE -In 5462..5495 -or
      $DEL_POSTCODE -In 5520..5540 -or
      $DEL_POSTCODE -eq 5560 -or
      $DEL_POSTCODE -eq 5600 -or
      $DEL_POSTCODE -In 5608..5609) {
        addfreight SAC $DEL_DELNO $del_dbtno
  }

  #SA D
  If ($DEL_POSTCODE -In 5220..5223 -or
      $DEL_POSTCODE -In 5434..5440 -or
      $DEL_POSTCODE -In 5601..5607 -or
      $DEL_POSTCODE -In 5630..5700) {
        addfreight SAD $DEL_DELNO $del_dbtno
  }

  If ($DEL_POSTCODE -eq 2880) {
    If ($DEL_DELADR2 -like "Broken Hill") {
        addfreight SAD $DEL_DELNO $del_dbtno
    }
  }

  #SA E
  If ($DEL_POSTCODE -In 5710..5734) {
    addfreight SAE $DEL_DELNO $del_dbtno
  }

  #DARWIN
  If ($DEL_POSTCODE -In 0800..0820 -or
      $DEL_POSTCODE -In 0828..0835) {
        addfreight DARWIN $DEL_DELNO $del_dbtno
  }

  #KATHRNE
  If ($DEL_POSTCODE -eq 0850 -or
      $DEL_POSTCODE -eq 0853) {
    addfreight KATHRNE $DEL_DELNO $del_dbtno
  }

  #TENCRK
  If ($DEL_POSTCODE -eq 0860) {
    addfreight TENCRK $DEL_DELNO $del_dbtno
  }


  #ALISPR
  If ($DEL_POSTCODE -eq 0870) {
    addfreight ALISPR $DEL_DELNO $del_dbtno
  }

  #NTOF
  If ($DEL_POSTCODE -In 0821..0822 -or
      $DEL_POSTCODE -In 0823..0827 -or
      $DEL_POSTCODE -In 0836..0849 -or
      $DEL_POSTCODE -In 0851..0852 -or
      $DEL_POSTCODE -In 0854..0859 -or
      $DEL_POSTCODE -In 0861..0869 -or
      $DEL_POSTCODE -In 0871..0900) {
    addfreight NTOF $DEL_DELNO $del_dbtno
  }

    #MELMET
  If ($DEL_POSTCODE -In 3000..3088 -or
      $DEL_POSTCODE -In 3101..3138 -or
      $DEL_POSTCODE -In 3140..3207 -or
      $DEL_POSTCODE -In 3802..3803 -or
      $DEL_POSTCODE -eq 3976 ) {
    addfreight MELMET $DEL_DELNO $del_dbtno
  }

  #MELOUT
  If ($DEL_POSTCODE -In 3089..3099 -or
      $DEL_POSTCODE -In 3765..3775 -or
      $DEL_POSTCODE -In 3785..3796 -or
      $DEL_POSTCODE -In 3804..3807 -or
      $DEL_POSTCODE -In 3910..3920 -or
      $DEL_POSTCODE -In 3926..3944 -or
      $DEL_POSTCODE -eq 3975 -or
      $DEL_POSTCODE -eq 3977) {
    addfreight MELOUT $DEL_DELNO $del_dbtno
  }

  #GEELONG
  If ($DEL_POSTCODE -In 3211..3220) {
    addfreight GEELONG $DEL_DELNO $del_dbtno
  }

  #BALRAT
  If ($DEL_POSTCODE -eq 3350) {
    addfreight BALRAT $DEL_DELNO $del_dbtno
  }

  #BNDGO
  If ($DEL_POSTCODE -eq 3550) {
    addfreight BNDGO $DEL_DELNO $del_dbtno
  }

  #TAS
  If ($DEL_POSTCODE -In 7000..7254 -or
      $DEL_POSTCODE -In 7258..7470) {
    addfreight TAS $DEL_DELNO $del_dbtno
  }

  #VICA
  If ($DEL_POSTCODE -eq 3139 -or
      $DEL_POSTCODE -In 3221..3230 -or
      $DEL_POSTCODE -In 3334..3345 -or
      $DEL_POSTCODE -In 3351..3373 -or
      $DEL_POSTCODE -In 3427..3465 -or
      $DEL_POSTCODE -In 3551..3559 -or
      $DEL_POSTCODE -In 3607..3616 -or
      $DEL_POSTCODE -In 3630..3631 -or
      $DEL_POSTCODE -In 3658..3670 -or
      $DEL_POSTCODE -In 3750..3764 -or
      $DEL_POSTCODE -In 3777..3783 -or
      $DEL_POSTCODE -In 3797..3799 -or
      $DEL_POSTCODE -In 3808..3844 -or
      $DEL_POSTCODE -In 3869..3871 -or
      $DEL_POSTCODE -In 3921..3925 -or
      $DEL_POSTCODE -In 3945..3959 -or
      $DEL_POSTCODE -In 3978..3996) {
    addfreight VICA $DEL_DELNO $del_dbtno
  }

  #VICB
  If ($DEL_POSTCODE -In 3231..3333 -or
      $DEL_POSTCODE -In 3375..3424 -or
      $DEL_POSTCODE -In 3467..3549 -or
      $DEL_POSTCODE -In 3561..3599 -or
      $DEL_POSTCODE -In 3617..3629 -or
      $DEL_POSTCODE -In 3633..3649 -or
      $DEL_POSTCODE -In 3672..3749 -or
      $DEL_POSTCODE -In 3847..3865 -or
      $DEL_POSTCODE -In 3873..3909 -or
      $DEL_POSTCODE -In 3960..3971) {
    addfreight VICB $DEL_DELNO $del_dbtno
  }

  #TASIS
  If ($DEL_POSTCODE -In 7255..7257) {
    addfreight TASIS $DEL_DELNO $del_dbtno
  }

  #PERMET
  If ($DEL_POSTCODE -In 6000..6027 -or
      $DEL_POSTCODE -In 6050..6066 -or
      $DEL_POSTCODE -In 6070..6081 -or
      $DEL_POSTCODE -eq 6090 -or
      $DEL_POSTCODE -In 6100..6121 -or
      $DEL_POSTCODE -In 6147..6171) {
    addfreight PERMET $DEL_DELNO $del_dbtno
  }

  #WAA
  If ($DEL_POSTCODE -In 6028..6049 -or
      $DEL_POSTCODE -In 6067..6069 -or
      $DEL_POSTCODE -In 6082..6089 -or
      $DEL_POSTCODE -In 6091..6099 -or
      $DEL_POSTCODE -In 6122..6146 -or
      $DEL_POSTCODE -In 6172..6640) {
    addfreight WAA $DEL_DELNO $del_dbtno
  }

  #WAB
  If ($DEL_POSTCODE -In 6641..6797) {
    addfreight WAB $DEL_DELNO $del_dbtno
  }

  #BRISBANE
  If ($DEL_POSTCODE -In 4000..4021 -or
      $DEL_POSTCODE -In 4026..4054 -or
      $DEL_POSTCODE -In 4059..4069 -or
      $DEL_POSTCODE -In 4072..4123 -or
      $DEL_POSTCODE -eq 4127 -or
      $DEL_POSTCODE -In 4151..4155 -or
      $DEL_POSTCODE -In 4157..4163 -or
      $DEL_POSTCODE -In 4169..4179 -or
      $DEL_POSTCODE -In 4500..4501

     ) {
    addfreight BRISBANE $DEL_DELNO $del_dbtno
  }

  #BRISOUT
  If ($DEL_POSTCODE -In 4022..4024 -or
      $DEL_POSTCODE -In 4055..4058 -or
      $DEL_POSTCODE -In 4070..4071 -or
      $DEL_POSTCODE -In 4124..4126 -or
      $DEL_POSTCODE -In 4128..4150 -or
      $DEL_POSTCODE -eq 4156 -or
      $DEL_POSTCODE -In 4164..4168 -or
      $DEL_POSTCODE -In 4180..4182 -or
      $DEL_POSTCODE -In 4300..4305 -or
      $DEL_POSTCODE -In 4502..4504 -or
      $DEL_POSTCODE -In 4508..4509
     ) {
   addfreight BRISOUT $DEL_DELNO $del_dbtno
  }

  #GOLCOAST
  If ($DEL_POSTCODE -In 4185..4230) {
    If ($DEL_DELADR2 -notlike 'Austinville' -or
        $DEL_DELADR2 -notlike 'Beechmont' -or
        $DEL_DELADR2 -notlike 'Binna Burra' -or
        $DEL_DELADR2 -notlike 'Lower Beechmont' -or
        $DEL_DELADR2 -notlike 'Natural Bridge' -or
        $DEL_DELADR2 -notlike 'Neranwood' -or
        $DEL_DELADR2 -notlike 'Numinbah Valley' -or
        $DEL_DELADR2 -notlike 'OReilly' -or
        $DEL_DELADR2 -notlike 'Rita Island' -or
        $DEL_DELADR2 -notlike 'Southern Lamington' -or
        $DEL_DELADR2 -notlike 'Springbrook') {
          addfreight GOLCOAST $DEL_DELNO $del_dbtno
        } else {
          addfreight C9 $DEL_DELNO $del_dbtno
        }
    }

  #SUNCOAST
  If ($DEL_POSTCODE -In 4505..4506 -or
      $DEL_POSTCODE -eq 4510 -or
      $DEL_POSTCODE -In 4516..4519 -or
      $DEL_POSTCODE -In 4550..4551 -or
      $DEL_POSTCODE -In 4553..4568 -or
      $DEL_POSTCODE -In 4572..4573 -or
      $DEL_POSTCODE -eq 4575) {
        addfreight SUNCOAST $DEL_DELNO $del_dbtno
  }

  #TOOWMBA
  If ($DEL_POSTCODE -eq 4350) {
    addfreight TOOWMBA $DEL_DELNO $del_dbtno
  }

  #MARYB
  If ($DEL_POSTCODE -In 4650..4655) {
    addfreight MARYB $DEL_DELNO $del_dbtno
  }

  #BUNDY
  If ($DEL_POSTCODE -eq 4670) {
    addfreight BUNDY $DEL_DELNO $del_dbtno
  }

  #ROCKY
  If ($DEL_POSTCODE -In 4700..4701) {
    addfreight ROCKY $DEL_DELNO $del_dbtno
  }

  #MACKAY
  If ($DEL_POSTCODE -eq 4740) {
    addfreight MACAKY $DEL_DELNO $del_dbtno
  }

  #TOWNVLE
  If ($DEL_POSTCODE -In 4810..4815 -or
      $DEL_POSTCODE -eq 4817) {
    addfreight TOWNVLE $DEL_DELNO $del_dbtno
  }

  #CAIRNS
  If ($DEL_POSTCODE -eq 4870) {
    addfreight CAIRNS $DEL_DELNO $del_dbtno
  }

  #C1
  If ($DEL_POSTCODE -In 4231..4299 -or
      $DEL_POSTCODE -In 4306..4349 -or
      $DEL_POSTCODE -eq 4507 -or
      $DEL_POSTCODE -In 4511..4515 -or
      $DEL_POSTCODE -In 4520..4549 -or
      $DEL_POSTCODE -eq 4552 -or
      $DEL_POSTCODE -eq 4569 -or
      $DEL_POSTCODE -eq 4574) {
      addfreight C1 $DEL_DELNO $del_dbtno
  }

  #C2
  If ($DEL_POSTCODE -In 4351..4416 -or
      $DEL_POSTCODE -In 4418..4420 -or
      $DEL_POSTCODE -In 4424..4428 -or
      $DEL_POSTCODE -In 4455..4470 -or
      $DEL_POSTCODE -In 4600..4615) {
    addfreight C2 $DEL_DELNO $del_dbtno
  }

  #C3
  If ($DEL_POSTCODE -In 4570..4571 -or
      $DEL_POSTCODE -In 4576..4599 -or
      $DEL_POSTCODE -In 4616..4649 -or
      $DEL_POSTCODE -In 4656..4669 -or
      $DEL_POSTCODE -In 4671..4676) {
    addfreight C3 $DEL_DELNO $del_dbtno
  }

  #C4
  If ($DEL_POSTCODE -In 4677..4694 -or
      $DEL_POSTCODE -eq 4703 -or
      $DEL_POSTCODE -In 4714..4722) {
    addfreight C4 $DEL_DELNO $del_dbtno
  }

  #C5
    If ($DEL_POSTCODE -In 4709..4713 -or
        $DEL_POSTCODE -eq 4737 -or
        $DEL_POSTCODE -In 4742..4799) {
      addfreight C5 $DEL_DELNO $del_dbtno
    }

    If ($DEL_POSTCODE -eq 4741) {
      If ($DEL_DELADR2 -like 'Coppabella' -or
          $DEL_DELADR2 -like 'Eton' -or
          $DEL_DELADR2 -like 'Eton North' -or
          $DEL_DELADR2 -like 'Farleigh' -or
          $DEL_DELADR2 -like 'Gargett' -or
          $DEL_DELADR2 -like 'Halliday Bay' -or
          $DEL_DELADR2 -like 'Kuttabul' -or
          $DEL_DELADR2 -like 'Mount Ossa' -or
          $DEL_DELADR2 -like 'Pinnacle' -or
          $DEL_DELADR2 -like 'Pleystowe' -or
          $DEL_DELADR2 -like 'Seaforth' -or
          $DEL_DELADR2 -like 'Yalboroo') {
            addfreight C5 $DEL_DELNO $del_dbtno
        } else {
           addfreight C9 $DEL_DELNO $del_dbtno
        }
    }

    #C6
    If ($DEL_POSTCODE -In 4800..4804 -or
        $DEL_POSTCODE -eq 4806 -or
        $DEL_POSTCODE -In 4808..4809 -or
        $DEL_POSTCODE -eq 4818 -or
        $DEL_POSTCODE -In 4849..4860) {
          addfreight C6 $DEL_DELNO $del_dbtno
    }

    If ($DEL_POSTCODE -eq 4805) {
      If ($DEL_DELADR2 -like 'Bowen' -or
          $DEL_DELADR2 -like 'Delta' -or
          $DEL_DELADR2 -like 'Merinda' -or
          $DEL_DELADR2 -like 'Wueens Beach' -or
          $DEL_DELADR2 -like 'Rose Bay') {
            addfreight C6 $DEL_DELNO $del_dbtno
          } else {
            addfreight C9 $DEL_DELNO $del_dbtno
          }
    }

    If ($DEL_POSTCODE -eq 4807) {
      If ($DEL_DELADR2 -notlike 'Austinville' -or
          $DEL_DELADR2 -notlike 'Beechmont' -or
          $DEL_DELADR2 -notlike 'Binna Burra' -or
          $DEL_DELADR2 -notlike 'Lower Beechmont' -or
          $DEL_DELADR2 -notlike 'Natural Bridge' -or
          $DEL_DELADR2 -notlike 'Neranwood' -or
          $DEL_DELADR2 -notlike 'Numinbah Valley' -or
          $DEL_DELADR2 -notlike 'OReilly' -or
          $DEL_DELADR2 -notlike 'Rite Island' -or
          $DEL_DELADR2 -notlike 'Southern Lamington' -or
          $DEL_DELADR2 -notlike 'Springbrook') {
            addfreight C6 $DEL_DELNO $del_dbtno
          } else {
            addfreight C9 $DEL_DELNO $del_dbtno
          }
      }
    #C7
    If ($DEL_POSTCODE -In 4861..4869 -or
        $DEL_POSTCODE -In 4878..4880 -or
        $DEL_POSTCODE -In 4882..4883 -or
        $DEL_POSTCODE -eq 4886) {
          addfreight C7 $DEL_DELNO $del_dbtno
    }

    If ($DEL_POSTCODE -eq 4885) {
      If ($DEL_DELADR2 -notlike 'Butchers Creek' -or
         $DEL_DELADR2 -notlike 'Glen Allyn') {
           addfreight C7 $DEL_DELNO $del_dbtno
         } else {
           addfreight C9 $DEL_DELNO $del_dbtno
         }
    }

    #C8
    If ($DEL_POSTCODE -In 4477..4479 -or
        $DEL_POSTCODE -eq 4702 -or
        $DEL_POSTCODE -In 4726..4725 -or
        $DEL_POSTCODE -In 4727..4736 -or
        $DEL_POSTCODE -eq 4816 -or
        $DEL_POSTCODE -In 4820..4825) {
         addfreight C8 $DEL_DELNO $del_dbtno
    }

    #C9
    If ($DEL_POSTCODE -eq 4025 -or
        $DEL_POSTCODE -In 4183..4184 -or
        $DEL_POSTCODE -eq 4417 -or
        $DEL_POSTCODE -In 4421..4423 -or
        $DEL_POSTCODE -In 4429..4454 -or
        $DEL_POSTCODE -In 4471..4476 -or
        $DEL_POSTCODE -In 4480..4499 -or
        $DEL_POSTCODE -In 4695..4699 -or
        $DEL_POSTCODE -In 4704..4708 -or
        $DEL_POSTCODE -eq 4726 -or
        $DEL_POSTCODE -In 4738..4739 -or
        $DEL_POSTCODE -eq 4819 -or
        $DEL_POSTCODE -In 4826..4848 -or
        $DEL_POSTCODE -In 4871..4877 -or
        $DEL_POSTCODE -eq 4881 -or
        $DEL_POSTCODE -eq 4884 -or
        $DEL_POSTCODE -In 4887..4999) {
          addfreight C9 $DEL_DELNO $del_dbtno
        }

}
