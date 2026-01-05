' DSG_v42.47.5 Last Updated 10/23/2025
OnAirV3 = "ON"
CG_TRIGGER_ON_TAKE = "ON"
data_feedback = "OFF"
headshot_path = "E:\EDRIVE\"
ignore_read = "false"
clear_data_tabs = "OFF"
dim data_NBA
dim data_MLB
dim data_NHL
dim shot_type
dim made_or_miss
dim superpage_requests
dim super_tab_group
dim on_air_tabs
dim dtvi_nap
dim OnAirV3_1_time

sub OnSocketDataReceived(data)
  if data_feedback = "ON" then msgbox "receiving " & data
  if TrioCmd("sock:socket_is_connected") then TrioCmd("sock:send_socket_data " & chr(42) & vbCrLf) ' dummy reply for SMT

  ''''OnAirV3 Addition'''''''''''''''''''''''''''''
    if left(data, 11) = "tvg_set_tab" then
      temp_arr = split(data, " ")
      if temp_arr(1) = temp_arr(2) then exit sub
    end if
  '''''''''''''''''''''''''''''''''''''''''''''''''

  ' ------ BEGIN SRDI LIVE INSERT ------
  if (CheckTvgLiveInsertRequest(data)) then exit sub
  ' ------ END SRDI LIVE INSERT --------
  if data = "R\1\\97" then exit sub
  if data = "R\1\__\\13" then exit sub
  if data = "R\1\____\\33" then exit sub
  if left(data, 10) = "R\1\ \ \ \" then exit sub

  'data = Replace(data, "__", " ") 'for NBA socket strings from DTVI
  'data = Replace(data, "~~", ",")
  data = Replace(data, "/NJD/NJD_", "/NJ/NJ_") ' fixing NHL headshot path
  data = Replace(data, "/LAK/LAK_", "/LA/LA_")
  data = Replace(data, "/TBL/TBL_", "/TB/TB_")
  data = Replace(data, "/SJS/SJS_", "/SJ/SJ_")
  data = Replace(data, "\NJD\", "\NJ\")        ' fixing TVI+ tricodes
  data = Replace(data, "\LAK\", "\LA\")
  data = Replace(data, "\TBL\", "\TB\")
  data = Replace(data, "\SJS\", "\SJ\")
  data = Replace(data, chr(195), chr(233))    ' Montreal accented e

  tab_match = ""
  temp = Left(data, 9)
  if temp = "R\1\21000" then tab_match = "HOME PL" ' player home
  if temp = "R\1\20000" then tab_match = "AWAY PL" ' player away
  if temp = "R\1\23000" then tab_match = "HOME TE" ' team home
  if temp = "R\1\22000" then tab_match = "AWAY TE" ' team away
  if Left(data, 25) = "tvg_set_tab PLAYER 1\HOME" then tab_match = "HOME PL"
  if Left(data, 25) = "tvg_set_tab PLAYER 1\AWAY" then tab_match = "AWAY PL"
  if Left(data, 23) = "tvg_set_tab TEAM 1\home" then tab_match = "HOME TE"
  if Left(data, 23) = "tvg_set_tab TEAM 1\away" then tab_match = "AWAY TE"

  if data_NHL="OFF" then data = UCase(data)


  if Right(tab_match, 2) = "PL" and Mid(data, 10, 2) = "00" then data = Left(data, 9) & "69" & Mid(data, 12) ' 00 jersey substitution

  if tab_match <> "" then
    data_arr = split(data, "\")
    'msgbox super_tab_group
    tab_group_arr = split(super_tab_group, ",")

    temp_jersey = data_arr(3) ' jersey #'s need to be in this field for all superpages
    'msgbox temp_jersey & " ?????"
    if Len(temp_jersey) = 1 then temp_jersey = "0" & temp_jersey  ' add leading zero to jersey for consistency
    tab_match = Replace(tab_match, "PL", temp_jersey)

    for i = 0 to ubound(tab_group_arr)
      tab_custom = TrioCmd("tabfield:get_custom_property " & tab_group_arr(i))
      custom_property_arr = split(tab_custom, " ")
      'msgbox tab_group_arr(i) & " working " & tab_custom & "    *****"

      for x = 0 to ubound(custom_property_arr) - 1
        if Right(custom_property_arr(x), 4) & " " & Left(custom_property_arr(x+1), 2) = tab_match then
          new_value = data_arr(custom_property_arr(x+2))
          'msgbox tab_match & " is in tab " & tab_group_arr(i) & " data=" & super_value

          if custom_property_arr(2) = "3" then
            if new_value = "69" then new_value = "00"  ' double 00 jersey substitution
          end if

          if IsNumeric(new_value) = true then ' numeric data that might have math
            current_value = TrioCmd("trio:get_global_variable " & tab_group_arr(i) & "value")
            if current_value = "" then
              current_value = new_value
              TrioCmd("trio:set_global_variable " & tab_group_arr(i) & "value " & current_value)
            else
          'msgbox tab_match & " is in tab " & tab_group_arr(i) & " data=" & current_value & " " & new_value & "        " &  Left(custom_property_arr(x), 1)

              if Left(custom_property_arr(x), 1) = "A" or Left(custom_property_arr(x), 1) = "H" then
                temp_operator = Left(custom_property_arr(x+3), 1)
                if temp_operator = "+" then current_value = cstr(cdbl(new_value) + cdbl(current_value))
                if temp_operator = "-" then current_value = cstr(cdbl(new_value) - cdbl(current_value))
                if temp_operator = "*" then current_value = cstr(cdbl(new_value) * cdbl(current_value))
                if temp_operator = "/" then current_value = cstr(cdbl(new_value) / cdbl(current_value))
              end if

              if Left(custom_property_arr(x), 1) = "+" then current_value = cstr(cdbl(current_value) + cdbl(new_value))
              if Left(custom_property_arr(x), 1) = "-" then current_value = cstr(cdbl(current_value) - cdbl(new_value))
              if Left(custom_property_arr(x), 1) = "*" then current_value = cstr(cdbl(current_value) * cdbl(new_value))
              if Left(custom_property_arr(x), 1) = "/" then current_value = cstr(cdbl(current_value) / cdbl(new_value))
              TrioCmd("trio:set_global_variable " & tab_group_arr(i) & "value " & current_value)
              if current_value < 0 then current_value = abs(current_value)
            end if
          else                                   ' non numeric data
            if custom_property_arr(2) = "12" then data_arr(custom_property_arr(2)) = Replace(data_arr(custom_property_arr(2)), chr(47), chr(92)) ' switching slash direction for headshot path
            current_value = new_value
          end if
          'msgbox current_total & " is current total"
        end if
      next

      prefix = ""
      suffix = ""
      temp_label = ""
      if InStr(tab_custom, "PREFIX=")  <> 0 then prefix = Mid(tab_custom, InStr(tab_custom, "PREFIX=") + 7)
      if InStr(prefix, "SUFFIX=")      <> 0 then prefix = Mid(prefix, 1, InStr(prefix, "SUFFIX=") - 2)
      if InStr(tab_custom, "SUFFIX=")  <> 0 then suffix = Mid(tab_custom, InStr(tab_custom, "SUFFIX=") + 7)
      if InStr(tab_custom, "LABEL=ON") <> 0 then temp_label = " " & data_arr(custom_property_arr(2) - 1)

      if InStr(tab_custom, "+CONSTANT=")  <> 0 then
        temp = Mid(tab_custom,InStr(tab_custom, "+CONSTANT=") + 10)
        if InStr(temp, " ") > 0 then temp = Left(temp, InStr(temp, " "))
        current_value = cdbl(current_value) + cdbl(temp)
      end if

      if InStr(tab_custom, "PERCENT=ON")   <> 0 then current_value = current_value * 100

      if IsNumeric(current_value) = true and InStr(current_value, ".") <> 0 and InStr(tab_custom, "RND=") <> 0 then
        temp_precision = Mid(tab_custom, InStr(tab_custom, "RND=") + 4, 1)
        current_value = current_value + .0001
        current_value = Round(cdbl(current_value), temp_precision)
      end if

      TrioCmd("page:set_property " & tab_group_arr(i) & " " & prefix & current_value & temp_label & suffix)
    next

    superpage_queue()
    exit sub
  end if

  if UCase(Left(data, Len("tvg_set_tab ClassicShotChart"))) = "TVG_SET_TAB CLASSICSHOTCHART"   then
    build_NBA_shotsSRDI(data)
    exit sub
  end if

  'Build NHL Shot Charts by sending data to seperate Sub
  if Left(data, 12) = "W\4999\94070" or Left(data, 12) = "W\4999\94076" then
    build_NHL_shots(data)
    exit sub
  end if

  'OnAir and SRDI data delivered to tabfields
  if (InStr(1, data, "TVG_SET_TAB")) or (InStr(1, data, "tvg_set_tab")) then

    if Ucase(left(data, 18)) <> "TVG_SET_TAB PLAYER" and Ucase(left(data, 16)) <> "TVG_SET_TAB TEAM" then
      cmd = LTrim(Mid(data, 12))
      idx = InStr(1, cmd, " ")
      tab = Left(cmd, idx - 1)
      value = LTrim(Mid(cmd, Len(tab) + 1))
      if InStr(tab, ".") > 0 then
        s_tab = Split(tab, ".")
        Select Case s_tab(1)
            Case "HL"
                tab = s_tab(0) & ".HL"
            Case Else
                tab = s_tab(0) & "." & LCase(s_tab(1))
        End Select
      end if

      if InStr(tab, "-") > 0 then
        s_tab2 = Split(tab, "-")
        'msgbox s_tab2(1)
        Select Case s_tab2(1)
            Case "NUMROWS"
                tab = s_tab2(0) & "-NumRows"
            Case "NumRows"
                tab = s_tab2(0) & "-NumRows"
            Case "POSROWS"
                tab = s_tab2(0) & "-PosRows"
            Case "ROWSPACING"
                tab = s_tab2(0) & "-RowSpacing"
            Case "TEXTPOS.position"
                tab = s_tab2(0) & "-TextPos.position"
            Case Else
                tab = s_tab2(0) & "-" & LCase(s_tab2(1))
        End Select
      end if

      TrioCmd("page:set_property " & tab & " " & Ucase(value))
      Call m_Replace_Ordinals(tab,value)

    end if
  end if

  if Left(data, 1) = "R" or Left(data, 1) = "W" then receive_SMT(data) 'Send SMT Data to SMT Sub

  superpage_queue()
end sub


sub OnRead(PageName)
  ' ------ BEGIN SRDI LIVE INSERT ------
  if (CheckTvgLiveInsertRead(PageName)) then exit sub
  ' ------ END SRDI LIVE INSERT --------
  TrioCmd("sock:set_socket_data_separator " & vbCrLf)
  template = TrioCmd("page:getpagetemplate")
  send_page = PageName
  page_custom = TrioCmd("tabfield:get_custom_property A")
  page_custom_odds = TrioCmd("tabfield:get_custom_property A*")
  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  superpage_requests = ""
  OnAirV3_1_time  = true

  'msgbox send_page
  if data_NBA = "" and data_MLB = "" and data_NHL = "" then ' reset data variables if showscript recompiled
    TrioCmd("vtwtemplate:run_vtw_script reset_data_league")
  end if

  if Len(PageName) > 6 and IsNumeric(PageName) = true then RefreshSRDLTricodes(PageName)

  if ignore_read = "true" then 'for pushed pages
    ignore_read = "false"
    exit sub
  end if

  on_air_tabs = "["
  supertabs = ""
  max_custom_property = -1

  for each tab in tab_arr ' Custom property packaging for onair, superpage tabs, and smt pages
    flag = TrioCmd("tabfield:get_custom_property " & tab)  ' on next line... exclude all non-onair custom properties.
    if Left(flag, 4) <> "SMT=" and IsNumeric(flag) = false and flag <> "" and Ucase(Left(flag, 5)) <> "HOME " and Ucase(Left(flag, 5)) <> "AWAY " and Ucase(Left(flag, 7)) <> "TEAM XX" and Ucase(Left(flag, 9)) <> "PLAYER XX" and left(flag, 1) <> "x" and left(flag, 12) <> "OddsEventID=" then
      if OnAirV3 <> "ON" or InStr(flag, "{{") > 0 or InStr(flag, "MULTI[") > 0 then
        oa_tab = "['" & tab & "','" & flag & "'], "
        on_air_tabs = on_air_tabs & oa_tab
        if clear_data_tabs = "ON" then TrioCmd("page:set_property " & tab & "")
      end if
    elseif left(flag, 12) = "OddsEventID=" then
      parts = Split(flag, "=")
      If UBound(parts) > 0 Then
        eventID = Split(parts(1), ":")(0)
        TrioCmd("vtwtemplate:run_vtw_script refresh_odds_data " & eventID & "|" & tab)
      end if
    elseif Ucase(Left(flag, 5)) = "HOME " or Ucase(Left(flag, 5)) = "AWAY " then ' Superpage tabs and SMT request queue
      TrioCmd("trio:set_global_variable " & tab & "value " & "")
      if Mid(flag, 7, 1) = " " then flag = Left(flag, 5) & "0" & Mid(flag, 6) ' adding in jersey leading 0 for page match comparison

      superpage_candidate = Left(flag, 7)

      if InStr(superpage_requests, superpage_candidate) = 0 then ' new superpage and single tab added to list
        superpage_requests = superpage_requests & "_" & superpage_candidate & ":" & tab
      elseif InStr(superpage_requests, superpage_candidate) > 1 then ' add tabs to existing superpage
        temp_insert = InStr(superpage_requests, superpage_candidate)
        superpage_requests = Mid(superpage_requests, 1, temp_insert + 7) & tab & "," & Mid(superpage_requests, temp_insert + 8)
      end if

      custom_prop_arr = split(flag, " ")
      for i = 0 to ubound(custom_prop_arr) - 1 ' process within supertab custom property for more +- data request candidates
        superpage_candidate = ""
        temp_search = "+HOME-HOME*HOME/HOME+AWAY-AWAY*AWAY/AWAY"
        if InStr(temp_search, custom_prop_arr(i)) <> 0 and custom_prop_arr(i) <> "" then
          superpage_candidate = Right(custom_prop_arr(i), 4) & " " & Left(custom_prop_arr(i+1), 2)
        end if
        if InStr(superpage_requests, superpage_candidate) = 0 then     ' new superpage and single tab added to list
          superpage_requests = superpage_requests & "_" & superpage_candidate & ":" & tab
        elseif InStr(superpage_requests, superpage_candidate) > 1 then ' add tabs to existing superpage
          temp_insert = InStr(superpage_requests, superpage_candidate)
          temp = Mid(superpage_requests, temp_insert + 8)

          if InStr(temp, "_") <> 0 then temp = Mid(temp, 1, InStr(temp, "_")) ' clip at _ if not last in list
          if InStr(temp, tab) = 0 then                                        'add tab if not already in sublist
            superpage_requests = Mid(superpage_requests, 1, temp_insert + 7) & tab & "," & Mid(superpage_requests, temp_insert + 8)
          end if
        end if
      next
      if clear_data_tabs = "ON" then TrioCmd("page:set_property " & tab & "")
    end if

    if IsNumeric(flag) then ' SMT tabs
      if cint(flag) > max_custom_property then max_custom_property = cint(flag)
      if clear_data_tabs = "ON" then TrioCmd("page:set_property " & tab & "")
    end if
  next

  on_air_tabs  = on_air_tabs & "]"
  superpage_requests = Mid(superpage_requests, 2) ' cleanup first blank value
  if data_feedback = "ON" and superpage_requests <> "" then msgbox "SuperpageRequest: " & chr(13) &  Replace(Replace(superpage_requests, ":", chr(13)), "_", chr(13) & chr(13)) ' superpage request order

  'dtvi_nap = 800
  'if Len(superpage_requests) > 36 then dtvi_nap = 2000

  ' Send a request the new format for on-air 2
  if superpage_requests = "" then
      if on_air_tabs <> "[]" then
        if data_feedback = "ON" then msgbox "OnAirMessageOutPRE: on_air_get message_number=" & send_page & " query=" & on_air_tabs & " message_context=" & TrioCmd("page:getpagedescription") & vbCrLf
        if TrioCmd("sock:socket_is_connected") then TrioCmd("sock:send_socket_data on_air_get message_number=" & send_page & " query=" & on_air_tabs & " message_context=" & TrioCmd("page:getpagedescription") & vbCrLf)
      end if
      OnAirV3_1_time = false
  end if

  ' no data for SMT, so send first superpage request to start sequence
  if max_custom_property = -1 then superpage_queue()

  'if Ucase(Left(page_custom_odds, 13)) = "ODDS:EVENTID=" then
      'AcustomArr = split(page_custom_odds, "=")
      'TrioCmd("vtwtemplate:run_vtw_script refresh_odds_data " & AcustomArr(1))
  'end if

  if data_MLB = "ON" or data_NBA = "ON" then
    if TrioCmd("sock:socket_is_connected") and data_feedback = "ON" then msgbox "tvg_viz_read " & template & " " & send_page & " " & tabs & vbCrLf                                      'send SRDI NBA or MLB only
    if TrioCmd("sock:socket_is_connected") then TrioCmd("sock:send_socket_data tvg_viz_read " & template & " " & send_page & " " & tabs & vbCrLf)
    'if data_MLB = "ON" then exit sub
    exit sub
  end if

  if TrioCmd("trio:get_global_variable dtvi_pull") <> "" then ' set in message from .vtw
    send_page = TrioCmd("trio:get_global_variable dtvi_pull")
    TrioCmd("trio:set_global_variable dtvi_pull")
  end if

  ' SMT= page data request redirection
  if InStr(TrioCmd("tabfield:get_custom_property A"), "SMT=") <> 0 then send_page = Replace(TrioCmd("tabfield:get_custom_property A"), "SMT=", "")

  if max_custom_property > -1 then send_socket send_page, max_custom_property

    if template = "SL_Odds" then
      if TrioCmd("tabfield:get_custom_property B0110 ") = "xSPREAD" then
        TrioCmd("trio:sleep 600")
        SpreadAwayHandicap = TrioCmd("page:get_property H0205")
        SpreadHomeHandicap = TrioCmd("page:get_property I0205")
        if CDbl(SpreadAwayHandicap) = CDbl(SpreadHomeHandicap) then
          TrioCmd("page:set_property H0199 0")
          TrioCmd("page:set_property I0199 1")
          TrioCmd("page:set_property H0205 [PK")
        elseif CDbl(SpreadAwayHandicap) > CDbl(SpreadHomeHandicap) then
          TrioCmd("page:set_property H0199 0")
          TrioCmd("page:set_property I0199 1")
          TrioCmd("page:set_property H0205 " & temp)
        elseif CDbl(SpreadAwayHandicap) < CDbl(SpreadHomeHandicap) then
          TrioCmd("page:set_property H0199 1")
          TrioCmd("page:set_property I0199 0")
          TrioCmd("page:set_property I0205 " & temp)
        end if
      end if
    end if
    if template = "CG_Ticker_2TeamNote" then
      CustomInfo = TrioCmd("tabfield:get_custom_property H0102 ")
      if Left(CustomInfo,11) = "xNBASPREAD:" then
        CustomInfo_arr = Split(CustomInfo, ":")
        SpreadAwayHandicap = CustomInfo_arr(1)
        SpreadHomeHandicap = CustomInfo_arr(2)
        OddsAwayNickname = CustomInfo_arr(3)
        OddsHomeNickname = CustomInfo_arr(4)
        if CDbl(SpreadAwayHandicap) = CDbl(SpreadHomeHandicap) then
          temp =  OddsAwayNickname & " AT " & OddsHomeNickname & "[(PK)]"
          TrioCmd("page:set_property H0101 " & temp)
          'temp = "OddsEventID=" & gameid & ":OddsAwayNickname at OddsHomeNickname [(PK)]"
          'TrioCmd("tabfield:set_custom_property H0101 " & chr(34) & temp & chr(34))
        elseif CDbl(SpreadAwayHandicap) > CDbl(SpreadHomeHandicap) then
          temp =  OddsAwayNickname & " AT " & OddsHomeNickname & " [(" & SpreadHomeHandicap.UTF8Text & ")"
          TrioCmd("page:set_property H0101 " & temp)
          'temp = "OddsEventID=" & gameid & ":OddsAwayNickname at OddsHomeNickname [(SpreadHomeHandicap)]"
          'TrioCmd("tabfield:set_custom_property H0101 " & chr(34) & temp & chr(34))
        elseif CDbl(SpreadAwayHandicap) < CDbl(SpreadHomeHandicap) then
          temp =  OddsAwayNickname & " [(" & SpreadAwayHandicap & ")]" & " AT " & OddsHomeNickname
          TrioCmd("page:set_property H0101 " & temp)
          'temp = "OddsEventID=" & gameid & ":OddsAwayNickname [(SpreadAwayHandicap)] at OddsHomeNickname"
          'TrioCmd("tabfield:set_custom_property H0101 " & chr(34) & temp & chr(34))
        end if
      end if
    end if
end sub


sub send_socket(send_page, max_custom_property)
  'send_page & " is the request"
  II_command = "X" & chr(92) & "1" & chr(92) & send_page & chr(92) & send_page & chr(92)
  for i = 0 to max_custom_property
    II_command = II_command & cstr(i) & chr(92)
  next
  if data_NHL = "ON" then
    if TrioCmd("sock:socket_is_connected") then
      if data_feedback = "ON" then msgbox "TVI+MessageOut: " & II_command & chr(92) & vbCrLf
      TrioCmd("sock:send_socket_data " & II_command & chr(92) & vbCrLf)
      TrioCmd("trio:sleep 50")
    end if
  elseif data_NBA = "ON" then
    if data_feedback = "ON" then msgbox "DTVIMessageOutToVTW: " & II_command & chr(92) & vbCrLf
    TrioCmd("vtwtemplate:run_vtw_script send_dtvi_socket " & II_command)
  end if
end sub


sub superpage_queue()

  if superpage_requests = "" then ' no superpages in queue, send onair if exists, exit
    'if OnAirV3 = "ON" then
      if on_air_tabs <> "[]" then
        if OnAirV3_1_time = true then
          send_page = TrioCmd("page:getpagename")
          if data_feedback = "ON" then msgbox "OnAirMessageOutPOST: on_air_get message_number=" & send_page & " query=" & on_air_tabs & " message_context=" & TrioCmd("page:getpagedescription") & vbCrLf
          if TrioCmd("sock:socket_is_connected") then TrioCmd("sock:send_socket_data on_air_get message_number=" & send_page & " query=" & on_air_tabs & " message_context=" & TrioCmd("page:getpagedescription") & vbCrLf)
          on_air_tabs = "[]"
          OnAirV3_1_time = false
        end if
      end if
    'end if
    exit sub
  end if

  superpage_requests_arr = split(superpage_requests, "_")
  super_page_1 = Left(superpage_requests_arr(0), 7)

  super_tab_group = Mid(superpage_requests, InStr(superpage_requests, ":") + 1)  ' find tabs associated with incoming superpage data
  if InStr(super_tab_group, "_") <> 0 then super_tab_group = Mid(super_tab_group, 1, InStr(super_tab_group, "_") - 1) ' clip to the _ if not last in list

  if InStr(superpage_requests, "_") <> 0 then ' remove first superpage from queue
    superpage_requests = Mid(superpage_requests, InStr(superpage_requests, "_") + 1)
  else                                        ' no _ delimiter for last page in queue
    superpage_requests = ""
  end if

  if super_page_1 = "HOME TE" then
    if data_NBA = "ON" then TrioCmd("sock:send_socket_data tvg_viz_read SRDI_team 9600 TEAM" & vbCrLf)
    if data_NHL = "ON" then send_socket "23000", 46
    if data_MLB = "ON" then TrioCmd("sock:send_socket_data tvg_viz_read SRDI_team 9600 TEAM" & vbCrLf)
  elseif super_page_1 = "AWAY TE" then
    if data_NBA = "ON" then TrioCmd("sock:send_socket_data tvg_viz_read SRDI_team 9500 TEAM" & vbCrLf)
    if data_NHL = "ON" then send_socket "22000", 46
    if data_MLB = "ON" then TrioCmd("sock:send_socket_data tvg_viz_read SRDI_team 9500 TEAM" & vbCrLf)
  else


    if data_NHL = "ON" then '210XX
      dtvi_player_page = "210"                                         ' Home superpage
      if Left(super_page_1, 4) = "AWAY" then dtvi_player_page = "200"  ' Away superpage
      dtvi_player_page = dtvi_player_page & Right(super_page_1, 2)
      send_socket dtvi_player_page, 49                                 ' TVI Player
    end if

    if data_MLB = "ON" or data_NBA = "ON" then '98XX
      srdi_player_page = "98"                                         ' Home superpage
      if Left(super_page_1, 4) = "AWAY" then srdi_player_page = "97"  ' Away superpage
      srdi_player_page = srdi_player_page & Right(super_page_1, 2)
      TrioCmd("sock:send_socket_data tvg_viz_read " & "SRDI_player" & " " & srdi_player_page & " PLAYER " & vbCrLf)
    end if

  end if
end sub


sub receive_SMT(data)
  TrioCmd("gui:status_message data received: " & data)
  data = replace(data, "UpdateScript", "")
  data_arr = Split(data, "\")
  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  data_arr(4) = replace(data_arr(4), "|", "\")


  if data_arr(0) = "W" then

    if data_NBA = "ON" then
      if Left(data, 6) = "R\1\ \" then exit sub
      if data_arr(2) = "0" then exit sub

      if CLng(data_arr(2)) > 9999 and cLng(data_arr(2)) < 21000 then data_arr(2) = CLng(data_arr(2)) + 60000   'intercept DTVI pages 10000-13999 and add 60000

      if data_arr(2) = "70003" or data_arr(2) = "70013" then   'Player Shot Chart fill in other info fields
        TrioCmd("page:set_property B0000 " & data_arr(6))
        TrioCmd("page:set_property H0021 " & data_arr(4))
        TrioCmd("page:set_property B0210 " & UCase(data_arr(5)))
        TrioCmd("page:set_property H0010 " & UCase(data_arr(8)))
        TrioCmd("page:set_property H0011 " & UCase(data_arr(9)))
        TrioCmd("page:set_property H0031 " & data_arr(7))
        TrioCmd("page:set_property H1110 " & data_arr(10))
        TrioCmd("page:set_property H1120 " & data_arr(11))
        exit sub
      end if

      if data_arr(2) = "70004" or data_arr(2) = "70014" then   'Team Shot Chart fill in other info fields
        TrioCmd("page:set_property H0088 " & "0")
        TrioCmd("page:set_property B0210 " & UCase(data_arr(5)))
        TrioCmd("page:set_property B0000 " & data_arr(6))
        TrioCmd("page:set_property H0011 " & UCase(data_arr(9)))
        TrioCmd("page:set_property H1110 " & data_arr(10))
        TrioCmd("page:set_property H1120 " & data_arr(11))
        exit sub
      end if

      if data_arr(2) = "70001" or data_arr(2) = "70011" or data_arr(2) = "70002" or data_arr(2) = "70012" then   'Player Hex Shot Chart fill in other info fields

      shots_arr = split(data, "|")

      for i = 6 to Ubound(shots_arr)
        shots_arr(i) = replace(shots_arr(i), "\", "")
        shot_cluster_arr = split(shots_arr(i), ",")
        shottype_arr = shottype_arr & "," & shot_cluster_arr(2)
      next
      if (InStr(1, shottype_arr, "2")) then
        TrioCmd("page:set_property H1000 " & data_arr(5))
      else
        TrioCmd("page:set_property H1000 " & "3-PT FG " & data_arr(5))
      end if
      TrioCmd("page:set_property B0000 " & data_arr(6))
      TrioCmd("page:set_property H1120 " & "")
      TrioCmd("page:set_property H0060 " & "")
      TrioCmd("page:set_property H1101 " & data_arr(18))
      TrioCmd("page:set_property H1110 " & data_arr(19))
      TrioCmd("page:set_property H1201 " & data_arr(12))
      TrioCmd("page:set_property H1210 " & data_arr(13) & "/" & data_arr(14))
      TrioCmd("page:set_property H1301 " & data_arr(15))
      TrioCmd("page:set_property H1310 " & data_arr(16) & "/"  & data_arr(17))
      end if

      if data_arr(2) = "70001" or data_arr(2) = "70011" then   'Player Hex Shot Chart
        TrioCmd("page:set_property H0000 " & "1")
        TrioCmd("page:set_property H0100 " & UCase(data_arr(8)))
        TrioCmd("page:set_property H0200 " & UCase(data_arr(9)))
        TrioCmd("page:set_property H0010 " & data_arr(4))
        TrioCmd("page:saveas 70004")
        exit sub
      end if

      if data_arr(2) = "70002" or data_arr(2) = "70012" then   'Team Hex Shot Chart
        TrioCmd("page:set_property H0000 " & "0")
        TrioCmd("page:set_property H0100 " & "")
        TrioCmd("page:set_property H0200 " & UCase(data_arr(9)))
        TrioCmd("page:saveas 70004")
        exit sub
      end if
    end if

    page_list = TrioCmd("show:get_pages")

    if InStr(page_list, data_arr(2)) <> 0 then
      ignore_read = "true"
      TrioCmd("page:read " & data_arr(2))
      'TrioCmd("trio:sleep 100")


      cg_pages = "60105 60205 60305 60405 67105 67205 67305 67405"
      pushed_page = TrioCmd("page:getpagename")
      if InStr(cg_pages, pushed_page) <> 0 then
        if pushed_page = "60105" or pushed_page = "67105" then data_arr(9) = data_arr(12) & " " & data_arr(9)
        if pushed_page = "60205" or pushed_page = "67205" then data_arr(9) = data_arr(14) & " " & data_arr(9)
        if pushed_page = "60305" or pushed_page = "67305" then data_arr(9) = data_arr(16) & " " & data_arr(9)
        if pushed_page = "60405" or pushed_page = "67405" then data_arr(9) = data_arr(18) & " " & data_arr(9)
      end if
    else
      exit sub
    end if
  end if

  if data_NBA = "ON" then                                  ' DTVI NBA tag offset
    NBA_ii_offset = 1
  else
    NBA_ii_offset = 0
  end if

  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)

  for j = 1 to Ubound(tab_arr)'1 - 1
    if data = "R\1\  \\" then exit for
    if data = "R\1\ \\" then exit for
    current_custom = TrioCmd("tabfield:get_custom_property " & tab_arr(j))
    if current_custom <> "" and Len(current_custom) <= 3 then
      II_value = cint(current_custom) + 2 - NBA_ii_offset

      if UCase(Left(data_arr(II_value), 5)) = "MONTR" then
          if Len(UCase(data_arr(II_value))) = 8 then
            data_arr(II_value) = "MONTR" & chr(201) & "AL"
          end if
      end if


      TrioCmd("page:set_property " & tab_arr(j) & " " & UCase(data_arr(II_value)))
    end if
  next

  if data_arr(0) = "W" and data_NBA = "ON" and TrioCmd("page:getpagename") > 900000 then  ' DTVI NBA clear Customs
    TrioCmd("trio:sleep 200")
    for each item in tab_arr
      if Len(TrioCmd("tabfield:get_custom_property " & item)) > 3 then
        TrioCmd("tabfield:set_custom_property " & item & " x" & TrioCmd("tabfield:get_custom_property " & item))
      end if
    next
  end if

 if TrioCmd("page:getpagename") = "95116" or TrioCmd("page:getpagename") = "95115" or TrioCmd("page:getpagename") = "9285" or TrioCmd("page:getpagename") = "9286" then'Remove 0's from OOTS FS Push pageS when schedule text is visible in Status Tab Field
   if TrioCmd("page:get_property H0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property H0116 " & "")
     TrioCmd("page:set_property H0126 " & "")
   end if
   if TrioCmd("page:get_property I0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property I0116 " & "")
     TrioCmd("page:set_property I0126 " & "")
   end if
     if TrioCmd("page:get_property J0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property J0116 " & "")
     TrioCmd("page:set_property J0126 " & "")
   end if
   if TrioCmd("page:get_property K0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property K0116 " & "")
     TrioCmd("page:set_property K0126 " & "")
   end if
   if TrioCmd("page:get_property L0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property L0116 " & "")
     TrioCmd("page:set_property L0126 " & "")
   end if
   if TrioCmd("page:get_property M0130 ") = "SCHEDULED" then
     TrioCmd("page:set_property M0116 " & "")
     TrioCmd("page:set_property M0126 " & "")
   end if

        'if TrioCmd("page:get_property N0130") = "Scheduled" then TrioCmd("page:set_property N0116 """) and TrioCmd("page:set_property N0126 """)
        'TrioCmd("page:set_property H1120 " & data_arr(6))
        'exit sub
 end if

 if data_NHL = "OFF" then exit sub

  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  for i = 0 to 10
    if InStr(tabs, "B000" & cstr(i)) <> 0 then                         ' Tricode
    if Left(TrioCmd("page:get_property B000" & cstr(i)), 12) = "C:/GRAPHICS/" then
      tab_value = TrioCmd("page:get_property B000" & cstr(i))
      tab_value = Replace(tab_value, "NJD", "NJ ")
      tab_value = Replace(tab_value, "N.J", "NJ ")
      tab_value = Replace(tab_value, "LAK", "LA ")
      tab_value = Replace(tab_value, "L.A", "LA ")
      tab_value = Replace(tab_value, "TBL", "TB ")
      tab_value = Replace(tab_value, "T.B", "TB ")
      tab_value = Replace(tab_value, "SJS", "SJ ")
      tab_value = Replace(tab_value, "S.J", "SJ ")
      TrioCmd("page:set_property B000" & cstr(i) & " " & Trim(Mid(tab_value, InStrRev(tab_value, "/") + 1, 3)))
    end if
      if Left(TrioCmd("page:get_property B000" & cstr(i)), 3) = "NJD" then TrioCmd("page:set_property B000" & cstr(i) & " NJ")
      if Left(TrioCmd("page:get_property B000" & cstr(i)), 3) = "LAK" then TrioCmd("page:set_property B000" & cstr(i) & " LA")
      if Left(TrioCmd("page:get_property B000" & cstr(i)), 3) = "TBL" then TrioCmd("page:set_property B000" & cstr(i) & " TB")
      if Left(TrioCmd("page:get_property B000" & cstr(i)), 3) = "SJS" then TrioCmd("page:set_property B000" & cstr(i) & " SJ")
    end if
  next


  if InStr(tabs, "H0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property H0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0000") & "\" & TrioCmd("page:get_property B0000") & "_"
    TrioCmd("page:set_property H0021 " & headshot_file & TrioCmd("page:get_property H0011") & "_" & TrioCmd("page:get_property H0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "I0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property I0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0001") & "\" & TrioCmd("page:get_property B0001") & "_"
    TrioCmd("page:set_property I0021 " & headshot_file & TrioCmd("page:get_property I0011") & "_" & TrioCmd("page:get_property I0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "J0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property J0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0002") & "\" & TrioCmd("page:get_property B0002") & "_"
    TrioCmd("page:set_property J0021 " & headshot_file & TrioCmd("page:get_property J0011") & "_" & TrioCmd("page:get_property J0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "K0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property K0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0003") & "\" & TrioCmd("page:get_property B0003") & "_"
    TrioCmd("page:set_property K0021 " & headshot_file & TrioCmd("page:get_property K0011") & "_" & TrioCmd("page:get_property K0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "L0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property L0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0004") & "\" & TrioCmd("page:get_property B0004") & "_"
    TrioCmd("page:set_property L0021 " & headshot_file & TrioCmd("page:get_property L0011") & "_" & TrioCmd("page:get_property L0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "H0014") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property H0014"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0000") & "\" & TrioCmd("page:get_property B0000") & "_"
    TrioCmd("page:set_property H0014 " & headshot_file & TrioCmd("page:get_property H0011") & "_" & TrioCmd("page:get_property H0010") & "_960.png")
  end if
  end if

  if InStr(tabs, "H0024") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property H0024"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0001") & "\" & TrioCmd("page:get_property B0001") & "_"
    TrioCmd("page:set_property H0024 " & headshot_file & TrioCmd("page:get_property H0021") & "_" & TrioCmd("page:get_property H0020") & "_960.png")
  end if
  end if

  if InStr(tabs, "M0021") <> 0 then                            ' Player headshot
  if Left(TrioCmd("page:get_property M0021"), 12) = "C:/GRAPHICS/" then
    headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0005") & "\" & TrioCmd("page:get_property B0005") & "_"
    TrioCmd("page:set_property M0021 " & headshot_file & TrioCmd("page:get_property M0011") & "_" & TrioCmd("page:get_property M0010") & "_960.png")
  end if
  end if

if TrioCmd("page:getpagetemplate") = "FS_Stnd" or TrioCmd("page:getpagetemplate") = "FS_Stnd18Teams" then  ' Standings subsitutions


if TrioCmd("page:getpagetemplate") = "FS_Stnd" then
  for i = 0 to 8
    TrioCmd("page:set_property B000" & cstr(i) & " " & TrioCmd("page:get_property H0" & cstr(i+1) & "10"))
  next
    TrioCmd("page:set_property B0009" & " " & TrioCmd("page:get_property H1010"))
end if

if TrioCmd("page:getpagetemplate") = "FS_Stnd18Teams" then
    for i = 1 to 9
      B_string = "B000"
      H_string = "H0"
      TrioCmd("page:set_property " & B_String & cstr(i) & " " & TrioCmd("page:get_property " & H_String & cstr(i) & "10"))
    next
    for i = 10 to 16
      B_string = "B00" & cstr(i)
      H_string = "H"
      TrioCmd("page:set_property " & B_String & " " & TrioCmd("page:get_property " & H_String & cstr(i) & "10"))
    next
end if

if TrioCmd("page:getpagetemplate") = "FS_Stnd" then starting_count = 0
if TrioCmd("page:getpagetemplate") = "FS_Stnd18Teams" then starting_count = 1
if TrioCmd("page:getpagetemplate") = "FS_Stnd" then team_count = 9
if TrioCmd("page:getpagetemplate") = "FS_Stnd18Teams" then team_count = 16
      for i = starting_count to team_count
      B_string = "B000"
      if i > 9 then B_string = "B00"

      if TrioCmd("page:get_property " & B_string & cstr(i)) = "TORONTO"      then TrioCmd("page:set_property " & B_string & cstr(i) & " TOR")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "EDMONTON"     then TrioCmd("page:set_property " & B_string & cstr(i) & " EDM")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "WINNIPEG"     then TrioCmd("page:set_property " & B_string & cstr(i) & " WPG")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "MONTR" & chr(201) & "AL"     then TrioCmd("page:set_property " & B_string & cstr(i) & " MTL")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "VANCOUVER"    then TrioCmd("page:set_property " & B_string & cstr(i) & " VAN")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "CALGARY"      then TrioCmd("page:set_property " & B_string & cstr(i) & " CGY")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "OTTAWA"       then TrioCmd("page:set_property " & B_string & cstr(i) & " OTT")

      if TrioCmd("page:get_property " & B_string & cstr(i)) = "NY ISLANDERS" then TrioCmd("page:set_property " & B_string & cstr(i) & " NYI")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "WASHINGTON"   then TrioCmd("page:set_property " & B_string & cstr(i) & " WSH")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "PITTSBURGH"   then TrioCmd("page:set_property " & B_string & cstr(i) & " PIT")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "BOSTON"       then TrioCmd("page:set_property " & B_string & cstr(i) & " BOS")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "PHILADELPHIA" then TrioCmd("page:set_property " & B_string & cstr(i) & " PHI")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "NY RANGERS"   then TrioCmd("page:set_property " & B_string & cstr(i) & " NYR")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "NEW JERSEY"   then TrioCmd("page:set_property " & B_string & cstr(i) & " NJ")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "BUFFALO"      then TrioCmd("page:set_property " & B_string & cstr(i) & " BUF")

      if TrioCmd("page:get_property " & B_string & cstr(i)) = "TAMPA BAY"    then TrioCmd("page:set_property " & B_string & cstr(i) & " TB")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "CAROLINA"     then TrioCmd("page:set_property " & B_string & cstr(i) & " CAR")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "FLORIDA"      then TrioCmd("page:set_property " & B_string & cstr(i) & " FLA")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "CHICAGO"      then TrioCmd("page:set_property " & B_string & cstr(i) & " CHI")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "COLUMBUS"     then TrioCmd("page:set_property " & B_string & cstr(i) & " CBJ")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "NASHVILLE"    then TrioCmd("page:set_property " & B_string & cstr(i) & " NSH")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "DALLAS"       then TrioCmd("page:set_property " & B_string & cstr(i) & " DAL")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "DETROIT"      then TrioCmd("page:set_property " & B_string & cstr(i) & " DET")

      if TrioCmd("page:get_property " & B_string & cstr(i)) = "VEGAS"        then TrioCmd("page:set_property " & B_string & cstr(i) & " VGK")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "COLORADO"     then TrioCmd("page:set_property " & B_string & cstr(i) & " COL")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "MINNESOTA"    then TrioCmd("page:set_property " & B_string & cstr(i) & " MIN")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "ST. LOUIS"    then TrioCmd("page:set_property " & B_string & cstr(i) & " STL")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "ARIZONA"      then TrioCmd("page:set_property " & B_string & cstr(i) & " ARI")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "LOS ANGELES"  then TrioCmd("page:set_property " & B_string & cstr(i) & " LA")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "SAN JOSE"     then TrioCmd("page:set_property " & B_string & cstr(i) & " SJ")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "ANAHEIM"      then TrioCmd("page:set_property " & B_string & cstr(i) & " ANA")
      if TrioCmd("page:get_property " & B_string & cstr(i)) = "SEATTLE"      then TrioCmd("page:set_property " & B_string & cstr(i) & " SEA")
    next
  end if

  'if current page is 94000 or 94001,
  'get tab field H0051, Remove from H0052 and trim remaining, set
  'grab tab field I0051, Remove from I0052 and trim remaining

  if TrioCmd("page:getpagename") = "94001" then
     TrioCmd("page:set_property H0052 " & trim(replace(TrioCmd("page:get_property H0052"), TrioCmd("page:get_property H0051"), "")))
     TrioCmd("page:set_property I0052 " & trim(replace(TrioCmd("page:get_property I0052"), TrioCmd("page:get_property I0051"), "")))
  end if
    if TrioCmd("page:getpagename") = "94000" then
     TrioCmd("page:set_property H0052 " & trim(replace(TrioCmd("page:get_property H0052"), TrioCmd("page:get_property H0051"), "")))
     TrioCmd("page:set_property H0062 " & trim(replace(TrioCmd("page:get_property H0062"), TrioCmd("page:get_property H0061"), "")))
  end if

  pgNumber = TrioCmd("page:getpagename")

  if data_NHL = "ON" then
    if (pgNumber > 8100 and pgNumber < 8500) then
           TrioCmd("trio:sleep 50")
           Call m_NHLstatParser()
    end if
  end if

end sub

sub OnValueChanged(PageName, PropertyName, NewPropertyValue)

  if Len(PropertyName) <> 5 then exit sub
  if Left(PropertyName, 3) <> "B00" and PropertyName <> "A1000" then exit sub

  league = TrioCmd("page:get_property A0100") ' if league field is blank then use global value set in .vtw
  if league = "" or league = " " then league = TrioCmd("trio:get_global_variable league")

  path_image = "IMAGE*/_TeamElements/Logos/"
  path_geom = "GEOM*_TeamElements/Logos/"
  team_num = right(PropertyName, 2)
  tricode = NewPropertyValue
  tab_library = TrioCmd("page:get_property_keys")

  if Left(PropertyName, 1) = "A" then
    if InStr(tab_library, "A1000.tri-2D") <> 0 then TrioCmd("page:set_property A1000.tri-2D IMAGE*_TeamElements/Logos/" & league & "/2D/" & tricode)
  end if

  if Left(PropertyName, 1) = "B" then
    if InStr(tab_library, "B00" & team_num & ".tri-2D") <> 0 then TrioCmd("page:set_property B00" & team_num & ".tri-2D " & path_image & league & "/2D/" & tricode)
    if InStr(tab_library, "B00" & team_num & ".tri-Flat") <> 0 then TrioCmd("page:set_property B00" & team_num & ".tri-Flat " & path_image & league & "/Flat/" & tricode)
    if InStr(tab_library, "B00" & team_num & ".tri-3D") <> 0 then TrioCmd("page:set_property B00" & team_num & ".tri-3D " & path_geom & league & "/3D/" & tricode)
    if InStr(tab_library, "B00" & team_num & ".tri-Sign") <> 0 then TrioCmd("page:set_property B00" & team_num & ".tri-Sign " & path_image & league & "/Sign/" & tricode)
  end if

end sub


sub OnTake(PageName)
  if CG_TRIGGER_ON_TAKE <> "ON" then exit sub
  if PageName = cstr("99999999") then exit sub
  if Left(TrioCmd("page:getpagetemplate"), 10) = "CG_Ticker_" then
    TrioCmd("vtwtemplate:run_vtw_script Trio_to_cg_link " & TrioCmd("page:getpagetemplate"))
  end if
end sub


sub OnTakeOut(PageName)
  if CG_TRIGGER_ON_TAKE <> "ON" then exit sub
  if PageName = cstr("99999999") then exit sub
  if Left(TrioCmd("page:getpagetemplate"), 10) = "CG_Ticker_" then
    TrioCmd("vtwtemplate:run_vtw_script cg_link_outClick null")
  end if
end sub


sub On_Air_refresh()
  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  new_on_air_tabs = "["

  for each tab in tab_arr
      flag = TrioCmd("tabfield:get_custom_property " & tab)
      if (flag) <> "" then
         new_oa_tab = "['" & tab & "','" & flag & "'], "
         new_on_air_tabs  = new_on_air_tabs & new_oa_tab
      end if
  next

  ' Close off our tuple and lists
  new_on_air_tabs  = new_on_air_tabs & "]"


  ' --- OnAir Requests ---
  ' Send a request the new format for on-air 2
  TrioCmd("sock:send_socket_data on_air_get message_number=" & TrioCmd("page:getpagename") & " query=" & new_on_air_tabs & " message_context=" & TrioCmd("page:getpagedescription") + vbCrLf)
end sub


sub build_NHL_shots(data)
  data = replace(data, "TBL", "TB")
  data = replace(data, "NJD", "NJ")
  data = replace(data, "LAK", "LA")
  data = replace(data, "SJS", "SJ")
  data_arr = Split(data, "\")
  tabs = TrioCmd("page:get_tabfield_names")
  tab_arr = Split(tabs)
  page_list = TrioCmd("show:get_pages")
  if data_arr(2) = "94076" then data_arr(2) = "94070"

  if InStr(page_list, data_arr(2)) <> 0 then
      TrioCmd("page:read " & data_arr(2))
          TrioCmd("page:read " & "94070")
      ignore_read = "true"

  ShotChartTemplate = TrioCmd("page:getpagetemplate")


  if Left(data, 12) = "W\4999\94070" then          'Player Shot Chart
    TrioCmd("page:set_property H0098 " & "2")
    TrioCmd("page:set_property H0010 " & UCase(data_arr(7)))
    TrioCmd("page:set_property H0011 " & UCase(data_arr(8)))
    TrioCmd("page:set_property H0021 " & data_arr(6))
    TrioCmd("page:set_property H0031 " & data_arr(3))
    TrioCmd("page:set_property H0041 " & data_arr(6))
    TrioCmd("page:set_property B0000 " & data_arr(9))
    if ShotChartTemplate = "FS_PlyrShotChartNHL-2023" then
      TrioCmd("page:set_property H0005 " & UCase(data_arr(4)))
    else
      TrioCmd("page:set_property B0200 " & UCase(data_arr(4)))
    end if
  end if
  if Left(data, 12) = "W\4999\94076" then          'Team Shot Chart
    TrioCmd("page:set_property H0098 " & "0")
    TrioCmd("page:set_property H0010 " & UCase(data_arr(4)))
    TrioCmd("page:set_property H0011 " & UCase(data_arr(5)))
    TrioCmd("page:set_property B0000 " & data_arr(7))
    if ShotChartTemplate = "FS_PlyrShotChartNHL-2023" then
      TrioCmd("page:set_property H0005 " & UCase(data_arr(3)))
    else
      TrioCmd("page:set_property B0200 " & UCase(data_arr(3)))
    end if
    TrioCmd("page:set_property H0031 " & "")
    TrioCmd("page:set_property H0041 " & "")
  end if

  TrioCmd("page:set_property H0115 0")
  TrioCmd("page:set_property H0215 0")
  TrioCmd("page:set_property H0315 0")
  TrioCmd("page:set_property H0415 0")
  TrioCmd("page:set_property H0515 0")
  TrioCmd("page:set_property G0100 " & "")
  TrioCmd("page:set_property G0150 " & "")
  TrioCmd("page:set_property G0200 " & "")
  TrioCmd("page:set_property G0250 " & "")
  TrioCmd("page:set_property G0300 " & "")
  TrioCmd("page:set_property G0350 " & "")
  TrioCmd("page:set_property G0400 " & "")
  TrioCmd("page:set_property G0450 " & "")
  headshot_file = headshot_path & TrioCmd("page:get_property A0100") & "\HEADSHOTS\" & TrioCmd("page:get_property B0000") & "\" & TrioCmd("page:get_property B0000") & "_"
  TrioCmd("page:set_property H0021 " & headshot_file & TrioCmd("page:get_property H0011") & "_" & TrioCmd("page:get_property H0010") & "_960.png")

  for w = 8 to ubound(data_arr)
    if data_arr(w) = "A" then TrioCmd("page:set_property H0315 " & data_arr(w+1))
    if data_arr(w) = "B" then TrioCmd("page:set_property H0215 " & data_arr(w+1))
    if data_arr(w) = "C" then TrioCmd("page:set_property H0515 " & data_arr(w+1))
    if data_arr(w) = "D" then TrioCmd("page:set_property H0415 " & data_arr(w+1))
  next
  TrioCmd("page:set_property H0115 " + cstr(cint(TrioCmd("page:get_property H0215 "))+ cint(TrioCmd("page:get_property H0515 ")) + cint(TrioCmd("page:get_property H0415 "))))

  for q = 1 to 4
    if q = 1 then shot_type_var = "\A\"
    if q = 2 then shot_type_var = "\B\"
    if q = 3 then shot_type_var = "\C\"
    if q = 4 then shot_type_var = "\D\"
    if q = 1 then x_shot_tabfield = "G0200"
    if q = 2 then x_shot_tabfield = "G0100"
    if q = 3 then x_shot_tabfield = "G0400"
    if q = 4 then x_shot_tabfield = "G0300"
    if q = 1 then y_shot_tabfield = "G0250"
    if q = 2 then y_shot_tabfield = "G0150"
    if q = 3 then y_shot_tabfield = "G0450"
    if q = 4 then y_shot_tabfield = "G0350"

    if InStrRev(data, shot_type_var) > 20 then  ' see if the shot type exists, skip first part of message
      shot_type = mid(data, InStrRev(data, shot_type_var), len(data))
      shot_arr = split(shot_type, "\")

      x_shot_package = ""
      y_shot_package = ""

      for x = 3 to shot_arr(2) + 2
        shot_pair_arr = split(shot_arr(x), ",")

        if ShotChartTemplate = "FS_PlyrShotChartNHL-2023" then
        '''''''''''Pack all shots on one side? why? push all shots to 1 side of the rink????
        if cdbl(shot_pair_arr(0)) > 0 then
          shot_pair_arr(0) = cstr(cdbl(shot_pair_arr(0))* -1)
          shot_pair_arr(1) = cstr(cdbl(shot_pair_arr(1))* -1)
        end if
        '''''''''''''''''''''''''''''''''''''''''''''''''''
          x_shot_package = x_shot_package & "," & shot_pair_arr(0)
          y_shot_package = y_shot_package & "," & cstr(cdbl(shot_pair_arr(1)*-1))
        else
       if TrioCmd("page:get_property H0001") = "0" then
         if cdbl(shot_pair_arr(0)) > 0 then
           shot_pair_arr(0) = cstr(cdbl(shot_pair_arr(0))* -1)
           shot_pair_arr(1) = cstr(cdbl(shot_pair_arr(1))* -1)
         end if
       end if
          x_shot_package = x_shot_package & "," & shot_pair_arr(0)
          y_shot_package = y_shot_package & "," & shot_pair_arr(1)
        end if

      next

      x_shot_package = Mid(x_shot_package, 2, 999)
      y_shot_package = Mid(y_shot_package, 2, 999)

      TrioCmd("page:set_property " & x_shot_tabfield & " " & x_shot_package)
      TrioCmd("page:set_property " & y_shot_tabfield & " " & y_shot_package)
    end if
   next
end if
end sub



 sub build_NBA_shots(data)
  ' NBA shot chart clear values
  if Left(data, 13) = "W\10005\10013" or Left(data, 13) = "W\10005\10003" or Left(data, 13) = "W\10005\10014" or Left(data, 13) = "W\10005\10004" or Left(data, 13) = "W\10005\10005" then  ' shot chart
    ignore_read = "true"
    TrioCmd("page:read 70003")
    for i = 201 to 399
      if i <> 300 then TrioCmd("page:set_property 0" & cstr(i) & ".active 0")
    next

    shots_arr = split(data, "|")

    made_tally = 200
    miss_tally = 300

    for i = 6 to Ubound(shots_arr)
      shots_arr(i) = replace(shots_arr(i), "\", "")
      if shots_arr(i) = "NULL" then
        msgbox "No Shots Attempted/Made By Selected Player/Team"
        exit sub
      end if
      shot_cluster_arr = split(shots_arr(i), ",")
     ' msgbox shot_cluster_arr(3)

      new_shot_x = cstr(cdbl(shot_cluster_arr(0)) * .62) - 0
      new_shot_y = cstr(cdbl(shot_cluster_arr(1)) * -.62) + 104

      if shot_cluster_arr(3) = "0" then
        miss_tally = miss_tally + 1
        TrioCmd("page:set_property 0" & cstr(miss_tally) & ".active 1")
        TrioCmd("page:set_property 0" & cstr(miss_tally) & ".position " & new_shot_x & " " & new_shot_y & " 0")
      else
        made_tally = made_tally + 1
        TrioCmd("page:set_property 0" & cstr(made_tally) & ".active 1")
        TrioCmd("page:set_property 0" & cstr(made_tally) & ".position " & new_shot_x & " " & new_shot_y & " 0")
      end if
    next
  end if

  if Left(data, 13) = "W\10000\10002" or Left(data, 13) = "W\10000\10012" or Left(data, 13) = "W\10000\10001" or Left(data, 13) = "W\10000\10011" then  ' shot chart
    ignore_read = "true"
    TrioCmd("page:read FS_Spec_ShotChartNBA")
    TrioCmd("sock:send_socket_data *" & vbcrlf)  'response
    shottype_arr = ""
    shots_arr = split(data, "|")

   for i = 6 to Ubound(shots_arr)

     shots_arr(i) = replace(shots_arr(i), "\", "")
       if shots_arr(i) = "NULL" then
         msgbox "No Shots Attempted/Made By Selected Player/Team"
         exit sub
       end if
     shot_cluster_arr = split(shots_arr(i), ",")
     shottype_arr = shottype_arr & "," & shot_cluster_arr(2)
    next
    TrioCmd("page:set_property A10000 " & data)
  end if
end sub


sub build_NBA_shotsSRDI(data)
    ignore_read = "true"

    for i = 201 to 399
        if i <> 300 then TrioCmd("page:set_property 0" & cstr(i) & ".active 0")
    next

    shots_arr = split(data, "|")
    made_tally = 200
    miss_tally = 300
    'msgbox "shots_arr 0=" & shots_arr(0)
    'msgbox "shots_arr 1=" & shots_arr(1)
    'msgbox "shots_arr 2=" & shots_arr(2)
    'msgbox "shots_arr 3=" & shots_arr(3)
    'msgbox "shots_arr 4=" & shots_arr(4)
    if shots_arr(1) = "?,?,?,?" then
        TrioCmd("gui:error_message ---->NO SHOTS TO DISPLAY ---->> PLEASE SELECT ANOTHER PLAYER/TEAM <<----")
        TrioCmd("gui:set_statusbar_color 255 0 0")
        exit sub
    end if

    for i = 1 to Ubound(shots_arr)
        shot_cluster_arr = split(shots_arr(i), ",")

        if shot_cluster_arr(0) = "?" then exit sub

        'if cdbl(shot_cluster_arr(0)) > 0.5 then
          'shot_cluster_arr(0) = cstr(1 - cdbl(shot_cluster_arr(0)))
          'shot_cluster_arr(1) = cstr(1 - cdbl(shot_cluster_arr(1)))
       'end if

        srdi_x = cdbl(shot_cluster_arr(1))
        srdi_y = cdbl(shot_cluster_arr(0))

        new_shot_x = cstr(srdi_x * -305 + 152) ' Adjusted scale for X
        new_shot_y = cstr((srdi_y) * -580 + 128) ' Adjusted scale for Y

        if Ucase(shot_cluster_arr(3)) = "MISS" then
            miss_tally = miss_tally + 1
            TrioCmd("page:set_property 0" & cstr(miss_tally) & ".active 1")
            TrioCmd("page:set_property 0" & cstr(miss_tally) & ".position " & new_shot_x & " " & new_shot_y & " 0")
        else
            made_tally = made_tally + 1
            TrioCmd("page:set_property 0" & cstr(made_tally) & ".active 1")
            TrioCmd("page:set_property 0" & cstr(made_tally) & ".position " & new_shot_x & " " & new_shot_y & " 0")
        end if
    next
end sub


Sub RefreshSRDLTricodes(PageName)
   ' refresh tricodes for SRDL pages
    league = TrioCmd("page:get_property A0100")
    path_image = "IMAGE*/_TeamElements/Logos/"
    path_geom = "GEOM*_TeamElements/Logos/"
    full_tab_library = TrioCmd("page:get_property_keys")

    if InStr(full_tab_library, "A1000.tri-2D") <> 0 then
      current_tricode = TrioCmd("page:get_property A1000")
      TrioCmd("page:set_property A1000.tri-2D " & path_image & league & "/2D/" & current_tricode & " ")
    end if

    for i = 0 to 9
      if InStr(full_tab_library, "B000" & cstr(i) & ".tri-2D") <> 0 then
        current_tricode = TrioCmd("page:get_property B000" & cstr(i))
        TrioCmd("page:set_property B000" & cstr(i) & ".tri-2D " & path_image & league & "/2D/" & current_tricode & " ")
      end if
    next

    for i = 10 to 99
      if InStr(full_tab_library, "B00" & cstr(i) & ".tri-2D") <> 0 then
        current_tricode = TrioCmd("page:get_property B00" & cstr(i))
        TrioCmd("page:set_property B00" & cstr(i) & ".tri-2D " & path_image & league & "/2D/" & current_tricode & " ")
      end if
    next

    for i = 0 to 9
      if InStr(full_tab_library, "B000" & cstr(i) & ".tri-3D") <> 0 then
        current_tricode = TrioCmd("page:get_property B000" & cstr(i))
        TrioCmd("page:set_property B000" & cstr(i) & ".tri-3D " & path_geom & league & "/3D/" & current_tricode & " ")
      end if
    next

    for i = 10 to 99
      if InStr(full_tab_library, "B00" & cstr(i) & ".tri-3D") <> 0 then
        current_tricode = TrioCmd("page:get_property B00" & cstr(i))
        TrioCmd("page:set_property B00" & cstr(i) & ".tri-3D " & path_geom & league & "/3D/" & current_tricode & " ")
      end if
    next

    for i = 0 to 9
      if InStr(full_tab_library, "B000" & cstr(i) & ".tri-Sign") <> 0 then
        current_tricode = TrioCmd("page:get_property B000" & cstr(i))
        TrioCmd("page:set_property B000" & cstr(i) & ".tri-Sign " & path_image & league & "/Sign/" & current_tricode & " ")
      end if
    next

        for i = 10 to 99
      if InStr(full_tab_library, "B00" & cstr(i) & ".tri-Sign") <> 0 then
        current_tricode = TrioCmd("page:get_property B00" & cstr(i))
        TrioCmd("page:set_property B00" & cstr(i) & ".tri-Sign " & path_image & league & "/Sign/" & current_tricode & " ")
      end if
    next
end sub


sub league_select(league)
  data_NBA = "OFF"
  data_NHL = "OFF"
  data_MLB = "OFF"
  if league = "NBA" then data_NBA = "ON"
  if league = "NHL" then data_NHL = "ON"
  if league = "MLB" then data_MLB = "ON"
  TrioCmd("trio:set_global_variable league " & league)
end sub

sub m_Replace_Ordinals(tab,value)
    TxtLine = Ucase(value)
    TxtLine = Replace(TxtLine, "[", "[ ")
    TxtLine = Replace(TxtLine, "]", " ]")
    TxtLine = Replace(TxtLine, "(", "( ")
    TxtLine = Replace(TxtLine, ")", " )")
    TxtLine = Replace(TxtLine, "-", "- ")

    NuVal = split(TxtLine, " ")

    for i = lbound(NuVal) to ubound(NuVal)
        if InStr(1, NuVal(i), "st", VBTextCompare) > 0 then
            if IsNumeric(Left(NuVal(i),1)) = true then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "ST", "st", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "nd", VBTextCompare) > 0 then
            if IsNumeric(Left(NuVal(i),1)) = true then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "ND", "nd", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "rd", VBTextCompare) > 0 then
            if IsNumeric(Left(NuVal(i),1)) = true then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "RD", "rd", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "th", VBTextCompare) > 0 then
            if IsNumeric(Left(NuVal(i),1)) = true then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "TH", "th", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "at", VBTextCompare) > 0 then
            if Len(NuVal(i)) = 2 then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "AT", "at", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "vs", VBTextCompare) > 0 then
            if Len(NuVal(i)) = 2 then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "VS", "vs", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "lhp", VBTextCompare) > 0 then
            if Len(NuVal(i)) = 3 then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "LHP", "lhp", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if

        if InStr(1, NuVal(i), "rhp", VBTextCompare) > 0 then
            if Len(NuVal(i)) = 3 then
                v1 = NuVal(i)
                TmpVal = Replace(v1, "RHP", "rhp", 1, 9, 1)
                NuVal(i) = TmpVal
            else
                TmpVal = NuVal(i)
            end if
        end if
    next

    for x = lbound(NuVal) to ubound(NuVal)
        tmpString = tmpString & NuVal(x) & " "
    next

    NewVal = Trim(tmpString)
    NewVal = Replace(NewVal, "[ ", "[")
    NewVal = Replace(NewVal, " ]", "]")
    NewVal = Replace(NewVal, "( ", "(")
    NewVal = Replace(NewVal, " )", ")")
    NewVal = Replace(NewVal, "- ", "-")
    NewVal = Trim(NewVal)

    TrioCmd("page:set_property " & tab & " " & NewVal)
end sub

Sub m_UnabbrevCGLink()
    on error resume next
    Name = TrioCmd("page:getpagetemplate")
    Select Case Name
        Case "CG_Ticker_PlayerStats", "CG_Ticker_TeamStats"
            For i = 1 to 4
                a_Category = TrioCmd("page:get_property " & "H0" & i & "00")

                Select Case a_Category
                    Case "POINTS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "PTS")
                                        Case "FG"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "FG")
                    Case "REBOUNDS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "REBS")
                    Case "ASSISTS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "ASTS")
                    Case "STEALS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "STLS")
                    Case "BLOCKS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "BLKS")
                    Case "MINUTES"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "MINS")
                    Case "3-PT FG"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "3-PT FG")
                    Case ""
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "")
                    Case Else
                        TrioCmd("page:set_property " & "H0" & i & "00 " & a_Category)
                End Select
            Next
        Case Else
            Tacos = "Good"
    End Select
End Sub

Sub m_UnabbrevL3rd()
    on error resume next
    Name = TrioCmd("page:getpagetemplate")
    Select Case Name
        Case "L3_TeamPlyrStat"
            For i = 1 to 5
                a_Category = TrioCmd("page:get_property " & "H0" & i & "00")

                Select Case a_Category
                    Case "POINTS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "PTS")
                                        Case "FG"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "FG")
                    Case "REBOUNDS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "REBS")
                    Case "ASSISTS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "ASTS")
                    Case "STEALS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "STLS")
                    Case "BLOCKS"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "BLKS")
                    Case "MINUTES"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "MINS")
                    Case "3-PT FG"
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "3-PT FG")
                    Case ""
                        TrioCmd("page:set_property " & "H0" & i & "00 " & "")
                    Case Else
                        TrioCmd("page:set_property " & "H0" & i & "00 " & a_Category)
                End Select
            Next
        Case Else
            Tacos = "Good"
    End Select
End Sub

sub m_NBAstatParser()
    TrioCmd("trio:sleep 50")
    on error resume next

    Name = TrioCmd("page:getpagetemplate")

    if Name = "CG_Ticker_PlayerStats" then
        a_statArray = TrioCmd("page:get_property D0001")
        s_statArray = split(a_statArray, "\")
    else
        a_statArray = TrioCmd("page:get_property Z*")
        s_statArray = split(a_statArray, "\")
    end if

    'Establish Game Time (Today or Tonight)
    a_cTime = Time
    s_ampm = Split(a_cTime, " ")
    s_cTime = Split(a_cTime, ":")

    if s_ampm(1) = "AM" then
        if s_cTime(0) = "12" then
            v_gameTime = "TONIGHT"
        elseif s_cTime(0) < 6 then
            v_gameTime = "TONIGHT"
        else
            v_gameTime = "TODAY"
        end if
    elseif s_ampm(1) = "PM" then
            if s_cTime(0) = "12" then
            v_gameTime = "TODAY"
        elseif s_cTime(0) > 4 then
            v_gameTime = "TONIGHT"
        else
            v_gameTime = "TODAY"
        end if
    end if

    'Split & Format Pulled Stat Array
    v_qual = UCase(s_statArray(4))
    v_pts = CDbl(s_statArray(8))
    v_rebs = CDbl(s_statArray(9))
    v_asts = CDbl(s_statArray(10))
    v_stls = CDbl(s_statArray(11))
    v_blks = CDbl(s_statArray(12))
    v_fg = s_statArray(13)
    v_3fg = s_statArray(14)
    v_ft = s_statArray(15)

    v_ptscat = s_statArray(16)
    v_rebscat = s_statArray(17)
    v_astscat = s_statArray(18)
    v_stlscat = s_statArray(19)
    v_blkscat = s_statArray(20)
    v_fgcat = s_statArray(21)
    v_3fgcat = s_statArray(22)
    v_ftcat = s_statArray(23)

    if v_qual = "SEASON" then
        v_fg = replace(v_fg, "%", "")
        v_fg = CInt(v_fg)

        if v_fg < 40 then
            v_fg = 0
        else
            v_fg = v_fg
        end if

        v_3fg = replace(v_3fg, "%", "")
        v_3fg = CInt(v_3fg)

        if v_3fg < 37 then
            v_3fg = 0
        else
            v_3fg = v_3fg
        end if
    else
        s_fg = split(v_fg, "/")
        a_fgm = s_fg(0)
        a_fga = s_fg(1)

        s_3fg = split(v_3fg, "/")
        a_3fgm = s_3fg(0)
        a_3fga = s_3fg(1)

        s_ft = split(v_ft, "/")
        a_ftm = s_ft(0)
        a_fta = s_ft(1)

        v_fga = CInt(a_fga)
        v_fgm = CInt(a_fgm)
        v_3fga = CInt(a_3fga)
        v_3fgm = CInt(a_3fgm)
        v_fta = CInt(a_fta)
        v_ftm = CInt(a_ftm)

        a_fgpct = Round((v_fgm)/(v_fga)*100)
        v_fgpct = CInt(a_fgpct)

        a_3fgpct = Round((v_3fgm)/(v_3fga)*100)
        v_3fgpct = CInt(a_3fgpct)

        if v_fga > 2 then
            if v_fgpct < 40 then
                v_fg = 0
            else
                v_fg = v_fgm & "/" & v_fga
            end if
        else
            v_fg = 0
        end if

        if v_3fga > 2 then
            if v_3fgpct < 40 then
                v_3fg = 0
            else
                v_3fg = v_3fgm & "/" & v_3fga
            end if
        else
            v_3fg = 0
        end if
    end if

    if (v_rebs < 2) then
        v_rebs = 0
    end if

    if (v_asts < 2) then
        v_asts = 0
    end if

    if (v_stls < 2) then
        v_stls = 0
    end if

    if (v_blks < 2) then
        v_blks = 0
    end if

    if (v_fg or v_3fg <> 0) then
        if (v_rebs and v_asts > 0) and (v_stls and v_blks > 0) then
            if (v_blks > v_rebs) or (v_stls > v_rebs) then
                v_rebs = 0
            elseif (v_blks > v_asts) or (v_stls > v_asts) then
                v_asts = 0
            else
                v_blks = 0
                v_stls = 0
            end if
        elseif (v_rebs and v_asts > 0) then
            if (v_rebs < 3) then
                v_rebs = 0
            end if

            if (v_asts < 3) then
                v_asts = 0
            end if
        end if
    end if

    'Establish PTS Display Value
    If v_pts = 0 Then
        a_DISpts = "0@POINTS"
    Elseif v_pts = 1 Then
        a_DISpts = v_pts & "@POINT"
    Elseif v_pts > 1 Then
        if right(s_statArray(8), 2) = ".0" then
            a_DISpts = v_pts & ".0@POINTS"
        else
            a_DISpts = v_pts & "@POINTS"
        end if
    end if

    'Establish REBS Display Value
    If v_rebs = 0 Then
        a_DISrebs = "0@REBOUNDS"
    Elseif v_rebs = 1 Then
        a_DISrebs = v_rebs & "@REBOUND"
    Elseif v_rebs > 1 Then
        if right(s_statArray(9), 2) = ".0" then
            a_DISrebs = v_rebs & ".0@REBOUNDS"
        else
            a_DISrebs = v_rebs & "@REBOUNDS"
        end if
    end if

    'Establish ASTS Display Value
    If v_asts = 0 Then
        a_DISasts = "0@ASSISTS"
    Elseif v_asts = 1 Then
        a_DISasts = v_asts & "@ASSIST"
    Elseif v_asts > 1 Then
        if right(s_statArray(10), 2) = ".0" then
            a_DISasts = v_asts & ".0@ASSISTS"
        else
            a_DISasts = v_asts & "@ASSISTS"
        end if
    end if

    'Establish STLS Display Value
    If v_stls = 0 Then
        a_DISstls = "0@STEALS"
    Elseif v_stls = 1 Then
        a_DISstls = v_stls & "@STEAL"
    Elseif v_stls > 1 Then
        if right(s_statArray(11), 2) = ".0" then
            a_DISstls = v_stls & ".0@STEALS"
        else
            a_DISstls = v_stls & "@STEALS"
        end if
    end if

    'Establish BLKS Display Value
    If v_blks = 0 Then
        a_DISblks = "0@BLOCKS"
    Elseif v_blks = 1 Then
        a_DISblks = v_blks & "@BLOCK"
    Elseif v_blks > 1 Then
        if right(s_statArray(12), 2) = ".0" then
            a_DISblks = v_blks & ".0@BLOCKS"
        else
            a_DISblks = v_blks & "@BLOCKS"
        end if
    end if

    'Establish FG Display Value
    if v_qual = "SEASON" then
        if v_fg = 0 then
            a_DISfg = "0@FG"
        else
            a_DISfg = v_fg & "%@FG"
        end if
    else
        If len(v_fg) < 2 Then
            a_DISfg = "0@FG"
        Else
            a_DISfg = v_fg & "@FG"
        end if
    end if

    'Establish 3FG Display Value
    if v_qual = "SEASON" then
        if v_3fg = 0 then
            a_DIS3fg = "0@3-PT FG"
        else
            a_DIS3fg = v_3fg & "%@3-PT FG"
        end if
    else
        If len(v_3fg) < 2 Then
            a_DIS3fg = "0@3-PT FG"
        Else
            a_DIS3fg = v_3fg & "@3-PT FG"
        end if
    end if

    'Establish FT Display Value
    If v_ft = 0 Then
        a_DISft = "0@FT"
    Else
        a_DISft = "0@FT"
    end if

    'String together and Apply
    if v_asts > v_rebs then
        a_DISarray = a_DISpts & ";" & a_DISasts & ";" & a_DISrebs & ";" & a_DISstls & ";" & a_DISblks & ";" & a_DISfg & ";" & a_DIS3fg & ";" & a_DISft
    else
        a_DISarray = a_DISpts & ";" & a_DISrebs & ";" & a_DISasts & ";" & a_DISstls & ";" & a_DISblks & ";" & a_DISfg & ";" & a_DIS3fg & ";" & a_DISft
    end if
    'msgbox a_DISarray
    'Split Parsed Array to Eliminate Stats
    s_DISarray = split(a_DISarray, ";")

    'Rebuild the whole string from the array parts.
    For x = 0 to UBound(s_DISarray)
        if x = 0 then
            tmpString = s_DISarray(0) & ";"
        else
            if left(s_DISarray(x),1) = "0" then
                tmpString = tmpString
            else
                tmpString = tmpString & s_DISarray(x) & ";"
            end if
        end if
    next
    'msgbox tmpString
    a_TRIMarray = Trim(tmpString)
    a_TRIMarrayLen = len(a_TRIMarray) - 1
    v_TRIMarray = left(a_TRIMarray, a_TRIMarrayLen)
    v_TRIMarrayRight = right(v_TRIMarray,1)
    v_TRIMarrayLen = len(v_TRIMarray) - 1

    if v_TRIMarrayRight = ";" then
        v_TRIMarray = left(v_TRIMarray, v_TRIMarrayLen)
    else
        v_TRIMarray = v_TRIMarray
    end if

    s_pts = split(a_DISpts, "@")

    if s_pts(0) = "99" then
        s_pts(0) = "0"
    else
        s_pts(0) = s_pts(0)
    end if

    s_statSplit = split(v_TRIMarray, ";")

    a_lastStat = UBound(s_statSplit)

    if a_lastStat = 0 then
        s_stat1 = split(s_statSplit(0), "@")
    elseif a_lastStat = 1 then
        s_stat1 = split(s_statSplit(0), "@")
        s_stat2 = split(s_statSplit(1), "@")
    elseif a_lastStat = 2 then
        s_stat1 = split(s_statSplit(0), "@")
        s_stat2 = split(s_statSplit(1), "@")
        s_stat3 = split(s_statSplit(2), "@")
    elseif a_lastStat = 3 then
        s_stat1 = split(s_statSplit(0), "@")
        s_stat2 = split(s_statSplit(1), "@")
        s_stat3 = split(s_statSplit(2), "@")
        s_stat4 = split(s_statSplit(3), "@")
    elseif a_lastStat > 3 then
        s_stat1 = split(s_statSplit(0), "@")
        s_stat2 = split(s_statSplit(1), "@")
        s_stat3 = split(s_statSplit(2), "@")
        s_stat4 = split(s_statSplit(3), "@")
        s_stat5 = split(s_statSplit(4), "@")
    end if

    Select Case Name
        Case "L3_TeamPlyrStat"
            if a_lastStat > 3 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & s_stat4(0))
                TrioCmd("page:set_property H0400 " & s_stat4(1))
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            elseif a_lastStat = 3 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & s_stat4(0))
                TrioCmd("page:set_property H0400 " & s_stat4(1))
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            elseif a_lastStat = 2 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            elseif a_lastStat = 1 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            elseif a_lastStat = 0 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & "")
                TrioCmd("page:set_property H0300 " & "")
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            elseif a_lastStat < 0 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & "")
                TrioCmd("page:set_property H0200 " & "")
                TrioCmd("page:set_property H0301 " & "")
                TrioCmd("page:set_property H0300 " & "")
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
            end if

            a_statChk1 = TrioCmd("page:get_property H0101")
            a_statChk2 = TrioCmd("page:get_property H0201")
            a_statChk3 = TrioCmd("page:get_property H0301")
            a_statChk4 = TrioCmd("page:get_property H0401")
            a_statChk5 = TrioCmd("page:get_property H0501")

            if len(a_statChk3) > 0 then
                Call m_UnabbrevL3rd()
            else
                Even_Austin = "LOVES TACOS"
            end if
        Case "CG_Ticker_PlayerStats"
            if a_lastStat > 3 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & s_stat4(0))
                TrioCmd("page:set_property H0400 " & s_stat4(1))
            elseif a_lastStat = 3 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & s_stat4(0))
                TrioCmd("page:set_property H0400 " & s_stat4(1))
            elseif a_lastStat = 2 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
            elseif a_lastStat = 1 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & s_stat3(0))
                TrioCmd("page:set_property H0300 " & s_stat3(1))
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
            elseif a_lastStat = 0 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & s_stat2(0))
                TrioCmd("page:set_property H0200 " & s_stat2(1))
                TrioCmd("page:set_property H0301 " & "")
                TrioCmd("page:set_property H0300 " & "")
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
            elseif a_lastStat < 0 then
                TrioCmd("page:set_property H0101 " & s_stat1(0))
                TrioCmd("page:set_property H0100 " & s_stat1(1))
                TrioCmd("page:set_property H0201 " & "")
                TrioCmd("page:set_property H0200 " & "")
                TrioCmd("page:set_property H0301 " & "")
                TrioCmd("page:set_property H0300 " & "")
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
            end if

            a_statChk1 = TrioCmd("page:get_property H0101")
            a_statChk2 = TrioCmd("page:get_property H0201")
            a_statChk3 = TrioCmd("page:get_property H0301")
            a_statChk4 = TrioCmd("page:get_property H0401")

            if len(a_statChk3) > 0 then
                Call m_UnabbrevCGLink()
            else
                Even_Austin = "LOVES TACOS"
            end if
        Case "TB_Main"
            if a_lastStat > 3 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1) & "  |  " & s_stat4(0) & " " & s_stat4(1))
            elseif a_lastStat = 3 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1) & "  |  " & s_stat4(0) & " " & s_stat4(1))
            elseif a_lastStat = 2 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1))
            elseif a_lastStat = 1 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_shifts(0) & " " & s_shifts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1))
            elseif a_lastStat = 0 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_fg(0) & " " & s_fg(1))
            elseif a_lastStat < 0 then
                TrioCmd("page:set_property H0210 " & s_pts(0) & " " & s_pts(1) & "  |  " & s_fg(0) & " " & s_fg(1))
            end if
        Case Else
            Tacos = "GOOD"
    End Select
end sub

sub m_NHLstatParser()
    TrioCmd("trio:sleep 50")
    'on error resume next

    Name = TrioCmd("page:getpagetemplate")
    a_statArray = TrioCmd("page:get_property Z*")
    if Left(a_statArray, 3) = " ; " then
       Exit Sub
    end if
    s_statArray = split(a_statArray, ";")
    a_zero = 0

    'Establish Game Time (Today or Tonight)
    a_cTime = Time
    s_ampm = Split(a_cTime, " ")
    s_cTime = Split(a_cTime, ":")

    if s_ampm(1) = "AM" then
        if s_cTime(0) = "12" then
            v_gameTime = "TONIGHT"
        elseif s_cTime(0) < 6 then
            v_gameTime = "TONIGHT"
        else
            v_gameTime = "TODAY"
        end if
    elseif s_ampm(1) = "PM" then
            if s_cTime(0) = "12" then
            v_gameTime = "TODAY"
        elseif s_cTime(0) > 4 then
            v_gameTime = "TONIGHT"
        else
            v_gameTime = "TODAY"
        end if
    end if

    'Split & Format Pulled Stat Array
    v_toi = s_statArray(0)
    v_shifts = CInt(s_statArray(1))
    v_goals = CInt(s_statArray(2))
    v_assists = CInt(s_statArray(3))
    v_shots = CInt(s_statArray(4))
    v_hits = CInt(s_statArray(5))
    v_blocks = CInt(s_statArray(6))
    v_attblocked = s_statArray(7)
    v_attmissed = s_statArray(8)
    v_faceoffwin = s_statArray(9)
    v_faceofftotal = s_statArray(10)

    if s_statArray(11) = " " then
        v_shotsagainst = 0
    elseif s_statArray(11) = "" then
        v_shotsagainst = 0
    else
        v_shotsagainst = CInt(s_statArray(11))
    end if
    v_saves = CInt(s_statArray(12))
    if s_statArray(13) = " " then
        v_savepct = 0
    elseif s_statArray(13) = "" then
        v_savepct = 0
    else
        v_savepct = s_statArray(13)
    end if
    v_rawrating = s_statArray(14)
    v_plyrpos = s_statArray(15)

    'Fix LW & RW Poisitions
    if v_plyrpos = "L" then
        v_plyrpos = "LW"
    elseif v_plyrpos = "R" then
        v_plyrpos = "RW"
    else
        v_plyrpos = v_plyrpos
    end if

    'Establish if Player is Skater or Goalie
    if v_savepct = 0 then
        'Calculate Shot Attempts
        if CInt(v_attblocked) = 0 and CInt(v_attmissed) = 0 then
            v_shotattempts = CInt(a_zero)
        else
            v_shotattempts = CInt(v_shots) + CInt(v_attblocked) + CInt(v_attmissed)
        end if
        'Determine if Rating is + or -
        if left(v_rawrating, 1) = "+" then
            v_rating = replace(v_rawrating, "+", "")
            v_rating = CInt(v_rating)
            if v_rating > 1 then
                v_rating = v_rating
            else
                v_rating = 0
            end if
        else
            v_rating = 0
        end if
        'Calculate Faceoff Win %
        if v_faceoffwin = " " then
            v_faceoffwin = 0
        end if

        if v_faceofftotal = " " then
            v_faceofftotal = 0
        end if

        If v_faceoffwin <> 0 Then
            v_faceoffpct = Round((v_faceoffwin)/(v_faceofftotal)*100)
        else
            v_faceoffpct = 0
        end if

        if v_plyrpos = "C" then
            if v_faceofftotal > 4 then
                if v_faceoffpct > 50 then
                    v_faceoffpct = v_faceoffpct
                else
                    v_faceoffpct = 0
                end if
            else
                v_faceoffpct = 0
            end if
        else
            v_faceoffpct = 0
        end if

        if (v_hits < 2) then
            v_hits = 0
        end if

        if (v_blocks < 2) then
            v_blocks = 0
        end if

        if (v_goals and v_assists > 0) then
            if (v_shotattempts = v_shots) then
                v_shotattempts = 0
            end if

            if (v_hits < 2) then
                v_hits = 0
            end if

            if (v_blocks < 2) then
                v_blocks = 0
            end if

            if (v_shotattempts > v_shots) then
                v_shots = 0
            end if

            if (v_shots = 0 and v_shotattempts > 0) then
                if (v_hits and v_blocks > 0) or (v_hits > 0 and v_blocks = 0) or (v_hits = 0 and v_blocks > 0) then
                    if (v_hits > v_blocks) then
                        v_blocks = 0
                    elseif (v_blocks > v_hits) or (v_blocks = v_hits) then
                        v_hits = 0
                    end if
                end if
            elseif (v_shots > 0 and v_shotattempts = 0) then
                if (v_hits and v_blocks > 0) or (v_hits > 0 and v_blocks = 0) or (v_hits = 0 and v_blocks > 0) then
                    if (v_hits > v_blocks) then
                        v_blocks = 0
                    elseif (v_blocks > v_hits) or (v_blocks = v_hits) then
                        v_hits = 0
                    end if
                end if
            end if
        else
            if (v_shotattempts = v_shots) then
                v_shotattempts = 0
            end if
            if (v_shotattempts > v_shots) then
                v_shots = 0
            end if
        end if

        'Establish TOI Display Value
        If len(v_toi) = 0 Then
            a_DIStoi = "0@TOI"
        Elseif len(v_toi) > 1 Then
            a_DIStoi = v_toi & "@TOI"
        end if

        'Establish Shifts Display Value
        If v_shifts = 0 Then
            a_DISshifts = "0@SHIFTS"
        Elseif v_shifts = 1 Then
            a_DISshifts = v_shifts & "@SHIFT"
        Elseif v_shifts > 1 Then
            a_DISshifts = v_shifts & "@SHIFTS"
        End if

        'Establish Goal Display Value
        If v_goals = 0 Then
            a_DISgoals = "0@GOALS"
        Elseif v_goals = 1 Then
            a_DISgoals = v_goals & "@GOAL"
        Elseif v_goals = 2 or v_goals > 3 Then
            a_DISgoals = v_goals & "@GOALS"
        Elseif v_goals = 3 Then
            a_DISgoals = v_goals & "@GOALS"
        end if

        'Establish Assist Display Value
        If v_assists = 0 Then
            a_DISassists = "0@ASSISTS"
        Elseif v_assists = 1 Then
            a_DISassists = v_assists & "@ASSIST"
        Elseif v_assists > 1 Then
            a_DISassists = v_assists & "@ASSISTS"
        end if

        'Establish Shot Display Value
        If v_shots = 0 Then
            a_DISshots = "0@SHOTS"
        Elseif v_shots = 1 Then
            a_DISshots = v_shots & "@SHOT"
        Elseif v_shots > 1 Then
            a_DISshots = v_shots & "@SHOTS"
        end if

        'Establish Hits Display Value
        If v_hits = 0 Then
            a_DIShits = "0@HITS"
        Elseif v_hits = 1 Then
            a_DIShits = v_hits & "@HIT"
        Elseif v_hits > 1 Then
            a_DIShits = v_hits & "@HITS"
        end if

        'Establish Blocks Display Value
        If v_blocks = 0 Then
            a_DISblocks = "0@BLOCKS"
        Elseif v_blocks = 1 Then
            a_DISblocks = v_blocks & "@BLOCKED SHOT"
        Elseif v_blocks > 1 Then
            a_DISblocks = v_blocks & "@BLOCKS"
        End if

        'Establish Shot Attempts Display Value
        If v_shotattempts = 0 Then
            a_DISshotattempts = "0@SHOT ATT."
        Elseif v_shotattempts = 1 Then
            a_DISshotattempts = v_shotattempts & "@SHOT ATT."
        Elseif v_shotattempts > 1 Then
            a_DISshotattempts = v_shotattempts & "@SHOT ATT."
        End if

        'Establish Saves Display Value
        If v_saves = 0 Then
            a_DISsaves = "0@SAVES"
        Elseif v_saves = 1 Then
            a_DISsaves = v_saves & "@SAVE"
        Elseif v_saves > 1 Then
            a_DISsaves = v_saves & "@SAVES"
        End if

        'Establish Shots Faced Display Value
        If v_shotsagainst = " " Then
            a_DISshotsagainst = "0@SHOTS FACED"
        Elseif v_shotsagainst = 1 Then
            a_DISshotsagainst = v_shotsagainst & "@SHOT FACED"
        Elseif v_shotsagainst > 1 Then
            a_DISshotsagainst = v_shotsagainst & "@SHOTS FACED"
        End if

        'Establish Saves Display Value
        If v_savepct = " " Then
            a_DISsavepct = "0@SAVE %"
        Elseif v_savepct = 1 Then
            a_DISsavepct = v_savepct & "@SAVE %"
        Elseif v_savepct > 1 Then
            a_DISsavepct = v_savepct & "@SAVE %"
        End if

        'Establish if Shot Attempts equal zero
        if v_shots = v_shotattempts then
            a_DISshotattempts = "0@SHOT ATT."
        end if

        'Establish FACEOFFS Display Value
        If v_faceoffpct = 0 Then
            a_DISfaceoffs = "0%@FACEOFFS"
        Else
            a_DISfaceoffs = v_faceoffpct & "%@FACEOFFS"
        end if

        'Establish Rating Display Value
        If v_rating = 0 Then
            a_DISrating = "0@RATING"
        Else
            a_DISrating = "+" & v_rating & "@RATING"
        End if

        'String together and Apply
        a_DISarray = a_DIStoi & ";" & a_DISshifts & ";" & a_DISgoals & ";" & a_DISassists & ";" & a_DIShits & ";" & a_DISblocks & ";" & a_DISshots & ";" & a_DISshotattempts & ";" & a_DISsaves & ";" & a_DISshotsagainst & ";" & a_DISsavepct & ";" & a_DISrating & ";" & a_DISfaceoffs

        'Split Parsed Array to Eliminate Stats
        s_DISarray = split(a_DISarray, ";")

        'Rebuild the whole string from the array parts.
        For x = 2 to UBound(s_DISarray)
            if left(s_DISarray(x),1) = "0" then
                tmpString = tmpString
            else
                tmpString = tmpString & s_DISarray(x) & ";"
            end if
        next

        a_TRIMarray = Trim(tmpString)
        a_TRIMarrayLen = len(a_TRIMarray)
        a_TRIMarrayLenMinus = (CInt(a_TRIMarrayLen) - 2)
        v_TRIMarray = left(a_TRIMarray, a_TRIMarrayLenMinus)
        v_TRIMarrayRight = right(v_TRIMarray, 1)
        v_TRIMarrayLen = len(v_TRIMarray)
        v_TRIMarrayLenMinus = (CInt(v_TRIMarrayLen) - 1)

        if v_TRIMarrayRight = ";" then
            v_TRIMarray = left(v_TRIMarray, v_TRIMarrayLenMinus)
        else
            v_TRIMarray = v_TRIMarray
        end if

        if InStr(v_TRIMarray, ";;") > 0 then
            a_doubleChk = "YES"
        else
            a_doubleChk = "NO"
        end if

        Do While a_doubleChk = "YES"
            v_TRIMarray = Replace(v_TRIMarray, ";;", ";")
            if InStr(v_TRIMarray, ";;") > 0 then
                a_doubleChk = "YES"
            else
                a_doubleChk = "NO"
            end if
        Loop

        v_TRIMarrayRightTest = right(v_TRIMarray, 5)
        v_TRIMarrayRightTest2 = right(v_TRIMarray, 7)

        if v_TRIMarrayRightTest = "RATIN" then
            v_TRIMarray = replace(v_TRIMarray, "RATIN", "RATING")
        elseif v_TRIMarrayRightTest2 = "FACEOFF" then
            v_TRIMarray = replace(v_TRIMarray, "FACEOFF", "FACEOFFS")
        end if

        s_toi = split(a_DIStoi, "@")
        s_shifts = split(a_DISshifts, "@")

        if InStr(v_TRIMarray, ";") > 0 then
            s_statSplit = split(v_TRIMarray, ";")
            a_lastStat = UBound(s_statSplit)
        else
            s_statSplit = v_TRIMarray
            a_lastStat = 0
        end if

        if a_lastStat = 0 then
            s_stat1 = split(s_statSplit, "@")
        elseif a_lastStat = 1 then
            s_stat1 = split(s_statSplit(0), "@")
            s_stat2 = split(s_statSplit(1), "@")
        elseif a_lastStat = 2 then
            s_stat1 = split(s_statSplit(0), "@")
            s_stat2 = split(s_statSplit(1), "@")
            s_stat3 = split(s_statSplit(2), "@")
        elseif a_lastStat = 3 then
            s_stat1 = split(s_statSplit(0), "@")
            s_stat2 = split(s_statSplit(1), "@")
            s_stat3 = split(s_statSplit(2), "@")
            s_stat4 = split(s_statSplit(3), "@")
        elseif a_lastStat > 3 then
            s_stat1 = split(s_statSplit(0), "@")
            s_stat2 = split(s_statSplit(1), "@")
            s_stat3 = split(s_statSplit(2), "@")
            s_stat4 = split(s_statSplit(3), "@")
            s_stat5 = split(s_statSplit(4), "@")
        end if

        Select Case Name
            Case "L3_TeamPlyrStat"
                TrioCmd("page:set_property H0101 " & s_toi(0))
                TrioCmd("page:set_property H0100 " & s_toi(1))
                if a_lastStat > 3 then
                    TrioCmd("page:set_property H0201 " & s_stat1(0))
                    TrioCmd("page:set_property H0200 " & s_stat1(1))
                    TrioCmd("page:set_property H0301 " & s_stat2(0))
                    TrioCmd("page:set_property H0300 " & s_stat2(1))
                    TrioCmd("page:set_property H0401 " & s_stat3(0))
                    TrioCmd("page:set_property H0400 " & s_stat3(1))
                    TrioCmd("page:set_property H0501 " & s_stat4(0))
                    TrioCmd("page:set_property H0500 " & s_stat4(1))
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                elseif a_lastStat = 3 then
                    TrioCmd("page:set_property H0201 " & s_stat1(0))
                    TrioCmd("page:set_property H0200 " & s_stat1(1))
                    TrioCmd("page:set_property H0301 " & s_stat2(0))
                    TrioCmd("page:set_property H0300 " & s_stat2(1))
                    TrioCmd("page:set_property H0401 " & s_stat3(0))
                    TrioCmd("page:set_property H0400 " & s_stat3(1))
                    TrioCmd("page:set_property H0501 " & s_stat4(0))
                    TrioCmd("page:set_property H0500 " & s_stat4(1))
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                elseif a_lastStat = 2 then
                    TrioCmd("page:set_property H0201 " & s_stat1(0))
                    TrioCmd("page:set_property H0200 " & s_stat1(1))
                    TrioCmd("page:set_property H0301 " & s_stat2(0))
                    TrioCmd("page:set_property H0300 " & s_stat2(1))
                    TrioCmd("page:set_property H0401 " & s_stat3(0))
                    TrioCmd("page:set_property H0400 " & s_stat3(1))
                    TrioCmd("page:set_property H0501 " & "")
                    TrioCmd("page:set_property H0500 " & "")
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                elseif a_lastStat = 1 then
                    TrioCmd("page:set_property H0201 " & s_shifts(0))
                    TrioCmd("page:set_property H0200 " & s_shifts(1))
                    TrioCmd("page:set_property H0301 " & s_stat1(0))
                    TrioCmd("page:set_property H0300 " & s_stat1(1))
                    TrioCmd("page:set_property H0401 " & s_stat2(0))
                    TrioCmd("page:set_property H0400 " & s_stat2(1))
                    TrioCmd("page:set_property H0501 " & "")
                    TrioCmd("page:set_property H0500 " & "")
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                elseif a_lastStat = 0 then
                    TrioCmd("page:set_property H0201 " & s_shifts(0))
                    TrioCmd("page:set_property H0200 " & s_shifts(1))
                    TrioCmd("page:set_property H0301 " & "0")
                    TrioCmd("page:set_property H0300 " & "SHOTS")
                    TrioCmd("page:set_property H0401 " & "")
                    TrioCmd("page:set_property H0400 " & "")
                    TrioCmd("page:set_property H0501 " & "")
                    TrioCmd("page:set_property H0500 " & "")
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                elseif a_lastStat < 0 then
                    TrioCmd("page:set_property H0201 " & s_shifts(0))
                    TrioCmd("page:set_property H0200 " & s_shifts(1))
                    TrioCmd("page:set_property H0301 " & "")
                    TrioCmd("page:set_property H0300 " & "")
                    TrioCmd("page:set_property H0401 " & "")
                    TrioCmd("page:set_property H0400 " & "")
                    TrioCmd("page:set_property H0501 " & "")
                    TrioCmd("page:set_property H0500 " & "")
                    TrioCmd("page:set_property H0601 " & "")
                    TrioCmd("page:set_property H0600 " & "")
                end if
                TrioCmd("page:set_property H0041 " & v_plyrpos)
                TrioCmd("page:set_property B0200 " & v_gameTime)
            Case "TB_Main"
                if a_lastStat > 3 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1) & "  |  " & s_stat4(0) & " " & s_stat4(1))
                elseif a_lastStat = 3 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1) & "  |  " & s_stat4(0) & " " & s_stat4(1))
                elseif a_lastStat = 2 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1) & "  |  " & s_stat3(0) & " " & s_stat3(1))
                elseif a_lastStat = 1 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_shifts(0) & " " & s_shifts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1) & "  |  " & s_stat2(0) & " " & s_stat2(1))
                elseif a_lastStat = 0 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_shifts(0) & " " & s_shifts(1) & "  |  " & s_stat1(0) & " " & s_stat1(1))
                elseif a_lastStat < 0 then
                    TrioCmd("page:set_property H0210 " & s_toi(0) & " " & s_toi(1) & "  |  " & s_shifts(0) & " " & s_shifts(1))
                end if
            Case Else
                Tacos = "GOOD"
        End Select
    else
        Select Case Name
            Case "L3_TeamPlyrStat"
                TrioCmd("page:set_property H0101 " & v_saves)
                TrioCmd("page:set_property H0100 " & "SAVES")
                TrioCmd("page:set_property H0201 " & v_shotsagainst)
                TrioCmd("page:set_property H0200 " & "SHOTS FACED")
                TrioCmd("page:set_property H0301 " & v_savepct)
                TrioCmd("page:set_property H0300 " & "SAVE %")
                TrioCmd("page:set_property H0401 " & "")
                TrioCmd("page:set_property H0400 " & "")
                TrioCmd("page:set_property H0501 " & "")
                TrioCmd("page:set_property H0500 " & "")
                TrioCmd("page:set_property H0601 " & "")
                TrioCmd("page:set_property H0600 " & "")
                TrioCmd("page:set_property H0041 " & v_plyrpos)
                TrioCmd("page:set_property B0200 " & v_gameTime)
            Case "TB_Main"
                TrioCmd("page:set_property H0210 " & v_saves & " SAVES ON " & v_shotsagainst & " SHOTS FACED  (" & v_savepct & " SAVE %)")
            Case Else
                Tacos = "GOOD"
        End Select
    end if
end sub


' ------ BEGIN SRDI LIVE INSERT ------
function CheckTvgLiveInsertRequest(data)
        'TrioCmd("gui:error_message " & data)

        dim header
        header = "tvg_live_insert_read"

        if (InStr(1, data, header)) then
                cmd = LTrim(Mid(data, Len(header) + 1))
                idx = InStr(1, cmd, " ")
                uid = Left(cmd, idx - 1)
                page = Replace(Trim(LTrim(Mid(cmd, Len(uid) + 1))), vbCrLf, "")
                'page = LTrim(Mid(cmd, Len(uid) + 1))

                TrioCmd("show:set_variable tvg_live_insert_id " & uid)
                TrioCmd("show:set_variable tvg_live_insert_page " & page)
                TrioCmd("page:read " & page)

                CheckTvgLiveInsertRequest = True
        else
                CheckTvgLiveInsertRequest = False
        end if
end function

function CheckTvgLiveInsertRead(page)
        dim uid
        uid = TrioCmd("show:get_variable tvg_live_insert_id")

        dim liveInsertPage
        liveInsertPage = TrioCmd("show:get_variable tvg_live_insert_page")

        if (liveInsertPage = page) then
                TrioCmd("sock:set_socket_data_separator " + vbCrLf)
                TrioCmd("sock:send_socket_data tvg_live_insert_read " & uid & " " & page & vbCrLf)
                TrioCmd("show:set_variable tvg_live_insert_page -")
                TrioCmd("show:set_variable tvg_live_insert_id -")

                CheckTvgLiveInsertRead = True
        else
                CheckTvgLiveInsertRead = False
        end if
end function
' ------ END SRDI LIVE INSERT --------
