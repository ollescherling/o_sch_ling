
#Persistent
#SingleInstance force

; Getting around the ctrl safe mode issue.
^b::  ; Ctrl + B
	Send !b  ; Send Alt + B
return

; -----------------------------
; Ctrl + B - Extract org numbers, titles, status, and main company
; -----------------------------
!b::
{
	Send {Ctrl up}
	Click
	Send ^a
	Sleep 200
	Send ^c
	Sleep 200
	Click
	Sleep 250

	if (!ErrorLevel) {
		text := clipboard
		Result := ""
		ResultEndedPerson := ""   ; ended, personnummer, NO U
		ResultEndedOrgU := ""     ; ended, orgnummer, WITH U

		; -------------------------
		; Logic for 'Org.nummer:'
		; -------------------------
		if InStr(text, "Org.nummer:") {
			UniqueNumbers := {}
			lines := StrSplit(text, "`n")
			mainOrg := ""
			endedEngagement := false

			if RegExMatch(text, "Org\.nummer:\s*(\d{6}-\d{4})", mainMatch)
				mainOrg := mainMatch1

			for index, line in lines {
				line := Trim(line)
				if (line = "")
					continue

				; Stop parsing
				if RegExMatch(line, "Ledamöter etc som lämnat bolaget de senaste 5 åren")
					break

				; Switch to ended-engagement mode
				if RegExMatch(line, "Engagemang som upphört enligt Bolagsverket") {
					endedEngagement := true
					continue
				}

				; Capture number
				if RegExMatch(line, "(\b\d{6}-\d{4}\b)", Match)
					number := Match1
				else
					continue

				; Detect personnummer (MM <= 31)
				mm := SubStr(number, 3, 2) + 0
				isPersonnummer := (mm <= 31)

				status := "-"
				if RegExMatch(line, "\bAvregistrerad\b")
					status := "Inactive"

				if RegExMatch(line, "(Ägare|Ordinarie ledamot|VD|Vice VD|Suppleant|Delägare|VVD och ord ledamot|VVD och suppleant|EVD|EVVD|Externa firmatecknare|Ord ledam o arbetstag repr|Suppl och arbetstagarrepr|Ord\. ledamot och arbetstagarrepr\.|Suppleant och arbetstagarrepr\.)", TitleMatch)
					title := TitleMatch1
				else
					continue

				mainFlag := (number = mainOrg) ? "Main" : "-"

				if (!UniqueNumbers.HasKey(number)) {
					UniqueNumbers[number] := 1

					; -------------------------
					; ACTIVE ENGAGEMENTS
					; -------------------------
					if (!endedEngagement) {
						row := number . "`t" . title . "`t" . status . "`t" . mainFlag . "`t-"
						if (Result != "")
							Result .= "`n"
						Result .= row
					}

					; -------------------------
					; ENDED ENGAGEMENTS
					; -------------------------
					else {
						if (isPersonnummer) {
							; NO U for personnummer
							row := number . "`t" . title . "`t" . status . "`t" . mainFlag . "`t-"
							if (ResultEndedPerson != "")
								ResultEndedPerson .= "`n"
							ResultEndedPerson .= row
						} else {
							; U only for orgnummer
							row := number . "`t" . title . "`t" . status . "`t" . mainFlag . "`tU"
							if (ResultEndedOrgU != "")
								ResultEndedOrgU .= "`n"
							ResultEndedOrgU .= row
						}
					}
				}
			}

			; -------------------------
			; Final append order
			; -------------------------
			if (ResultEndedPerson != "") {
				if (Result != "")
					Result .= "`n"
				Result .= ResultEndedPerson
			}

			if (ResultEndedOrgU != "") {
				if (Result != "")
					Result .= "`n"
				Result .= ResultEndedOrgU
			}
		}

		; -------------------------
		; Output to Excel
		; -------------------------
		if (Result != "") {
			clipboard := Result
			Run, "C:/Users/OLLESC/Desktop/Portfolio.xlsx"
			WinWaitActive, ahk_exe EXCEL.EXE, , 10
		} else {
			MsgBox, No organizational numbers found.
		}
	}
	else {
		MsgBox, No text in the clipboard.
	}
}
return

!x::
{
	ExitApp  ; Terminates the current AHK script
	return
}
