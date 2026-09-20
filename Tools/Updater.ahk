#Requires AutoHotkey v2.0
#SingleInstance Force

; #Include custom message box -- Required.
#Include "..\Includes\AcMsgBox.ahk"

/*
AutoCorrect2 Updater Tool
Author: kunkel321
Tool Used: Claude
Version: 6.1
Intended to use with AutoCorrect2 repo. Run the script, and it will download a temporary copy of the repository, then look for files that have been updated. The script makes repeat checks faster by saving revision-scoped state in Data\LastUpdateCheck.ini. PendingFiles records generic files still awaiting action for the stored commit SHA, including whether each pending item is a new file or an update, while UpdatedFiles records generic files successfully applied during that same update cycle with timestamps and actions. Updater.ahk/.exe use their dedicated [UpdaterState] section instead. A newly detected SHA starts a fresh cycle. It checks the latest GitHub commit SHA and uses an exact-SHA temporary source cache so every comparison and pending retry is tied to the precise repository revision recorded for that update cycle. The script then offers to replace local files with newer versions. Updates to existing files in the folders listed below as "RarelyUpdatedFolders" are unchecked by default because they are more likely to contain user customizations; new files remain checked by default. Any new files are listed separately.

AutoCorrectHotstrings.ahk is a special case and is never overwritten by default. Instead, the GitHub version is saved to the Core\ folder using the filename configured by [Files] NewTemporaryHotstrLib in acSettings.ini. The user can then use the UniqueStringExtractor to compare that upstream copy with their own library. If the user NEVER customizes AutoCorrectHotstrings.ahk and only wants to adopt the newer version, they can select the radio button to replace their existing library. The replacement option is shown in red as a warning.

Replace mode is treated as a compound operation: the upstream library is first copied to the active AutoCorrectHotstrings.ahk and then to the configured upstream snapshot, so the snapshot always represents the upstream version already processed.

Tools\Updater.ahk and Tools\Updater.exe are also special cases because this updater may be customized locally. Their active local copies are never overwritten automatically. Their persistent handled hashes and current-cycle candidate/status fields live in [UpdaterState]. Changed upstream copies are saved automatically under Tools\Updater_Upstream for comparison and are shown in read-only updater sections rather than selectable file lists.
*/

; --- CONFIG ---------------------------------------------------------------
; GitHub API endpoint for checking latest commit
GitHubApiUrl  := "https://api.github.com/repos/kunkel321/AutoCorrect2/commits?per_page=1"

; Base URL used to download an archive for one exact commit SHA.
RepositoryArchiveBaseUrl := "https://github.com/kunkel321/AutoCorrect2/archive"

; Optional: Enable debug logging to file
EnableDebugLog := 1

; Folders whose files will appear in update dialog but unchecked by default
RarelyUpdatedFolders := ["Data\"]

; Runtime-generated state/log files that must never be offered as repository updates.
ExcludedRuntimeFiles := [
    "Data\LastUpdateCheck.ini",
    "Data\ACLogContinuous.txt",
    "Data\AutoCorrectsLog.txt",
    "Data\ErrContextLog.txt",
    "Data\MCLogContinuous.txt",
    "Data\ManualCorrectionsLog.txt",
    "Debug\Updater_debug.txt",
    "Debug\ac2_error_debug_log.txt"
]
; -----------------------------------------------------------------------

; Setup debug logging
SplitPath(A_ScriptDir, , &parentDir)
debugDir := parentDir "\Debug"
debugFile := debugDir "\Updater_debug.txt"

LogDebug(msg) {
    if EnableDebugLog {
        ; Ensure Debug folder exists before attempting to write
        if !DirExist(debugDir) {
            try {
                DirCreate(debugDir)
            } catch Error as e {
                ; Silently fail if we can't create debug folder
                return
            }
        }
        
        timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
        try {
            FileAppend("[" timestamp "] " msg "`n", debugFile)
        } catch Error as e {
            ; Silently fail if we can't write debug log
        }
    }
}

CheckGitHubForUpdates() {
    ; Checks GitHub API for the latest commit info.
    ; Returns an object with {sha, date} on success, or error string on failure
    try {
        LogDebug("Checking GitHub API for latest commit...")
        
        ; Create HTTP request
        http := ComObject("MSXML2.XMLHTTP")
        http.Open("GET", GitHubApiUrl, false)
        http.SetRequestHeader("User-Agent", "AutoCorrect2-Updater")
        http.Send()
        
        if http.Status != 200 {
            return "GitHub API returned status " http.Status
        }
        
        ; Parse JSON response
        responseText := http.ResponseText
        LogDebug("GitHub API response received, parsing JSON...")
        LogDebug("Response length: " StrLen(responseText))
        LogDebug("First 500 chars: " SubStr(responseText, 1, 500))
        
        ; Extract sha and date using simple string parsing
        ; GitHub API format: "sha": "abcd1234...", "author": {"date": "2025-01-15T12:34:56Z"}
        sha := ""
        date := ""
        
        ; Extract SHA (commit hash)
        shaPos := InStr(responseText, '"sha"')
        LogDebug("shaPos: " shaPos)
        if shaPos > 0 {
            quotePos := InStr(responseText, '"',, shaPos + 5)
            LogDebug("First quote after sha: " quotePos)
            if quotePos > 0 {
                startPos := quotePos + 1
                endQuotePos := InStr(responseText, '"',, startPos)
                LogDebug("End quote: " endQuotePos)
                if endQuotePos > 0 {
                    sha := SubStr(responseText, startPos, endQuotePos - startPos)
                }
            }
        }
        
        ; Extract commit date
        ; JSON format: "date":"2025-11-17T..."
        datePos := InStr(responseText, '"date"')
        LogDebug("datePos: " datePos)
        if datePos > 0 {
            ; Find the colon after "date"
            colonPos := InStr(responseText, ":",, datePos)
            LogDebug("colonPos: " colonPos)
            ; Find the quote after the colon
            if colonPos > 0 {
                quotePos := InStr(responseText, '"',, colonPos)
                LogDebug("Quote after colon: " quotePos)
                if quotePos > 0 {
                    startPos := quotePos + 1
                    endQuotePos := InStr(responseText, '"',, startPos)
                    LogDebug("End quote: " endQuotePos)
                    if endQuotePos > 0 {
                        date := SubStr(responseText, startPos, endQuotePos - startPos)
                    }
                }
            }
        }
        
        LogDebug("Extracted SHA: " sha)
        LogDebug("Extracted date: " date)
        
        if sha = "" or date = "" {
            return "Could not parse GitHub API response"
        }
        
        LogDebug("Latest commit: SHA=" sha ", Date=" date)
        result := Map()
        result["sha"] := sha
        result["date"] := date
        return result
        
    } catch Error as e {
        return e.Message
    }
}

GetLastUpdateCheckInfo() {
    ; Read revision-scoped updater state from Data\LastUpdateCheck.ini.
    ; PendingFiles and UpdatedFiles describe only the stored commit SHA.
    checkFile := A_ScriptDir "\..\Data\LastUpdateCheck.ini"

    if !FileExist(checkFile) {
        LogDebug("LastUpdateCheck.ini not found - first run or data file missing")
        return Map()
    }

    try {
        lastSha := IniRead(checkFile, "UpdateCheck", "LastCommitHash", "")
        lastDate := IniRead(checkFile, "UpdateCheck", "LastCommitDate", "")

        if lastSha = "" or lastDate = "" {
            LogDebug("LastUpdateCheck.ini exists but is missing commit state")
            return Map()
        }

        pendingRaw := IniRead(checkFile, "UpdateCheck", "PendingFiles", "")
        pendingFiles := []
        if pendingRaw != "" {
            for item in StrSplit(pendingRaw, "|") {
                item := Trim(item)
                if item = ""
                    continue
                if IsExcludedPendingItem(item) {
                    LogDebug("Ignoring excluded pending entry: " item)
                    continue
                }
                pendingFiles.Push(item)
            }
        }

        updatedRaw := IniRead(checkFile, "UpdateCheck", "UpdatedFiles", "")
        updatedFiles := ParseUpdatedFileRecords(updatedRaw)

        LogDebug(
            "Last update check: SHA=" lastSha
            ", Date=" lastDate
            ", Pending=" pendingFiles.Length
            " file(s), Updated=" updatedFiles.Length " file(s)")

        result := Map()
        result["sha"] := lastSha
        result["date"] := lastDate
        result["pendingFiles"] := pendingFiles
        result["updatedFiles"] := updatedFiles
        return result

    } catch Error as e {
        LogDebug("Error reading LastUpdateCheck.ini: " e.Message)
        return Map()
    }
}

StoreUpdateCheckInfo(sha, date, pendingFiles := "", updatedFiles := "") {
    ; Store generic state for one upstream commit SHA. Updater.ahk/.exe special
    ; state is stored separately in [UpdaterState].
    checkFile := A_ScriptDir "\..\Data\LastUpdateCheck.ini"
    checkDir := A_ScriptDir "\..\Data"

    try {
        if !DirExist(checkDir)
            DirCreate(checkDir)

        IniWrite(sha, checkFile, "UpdateCheck", "LastCommitHash")
        IniWrite(date, checkFile, "UpdateCheck", "LastCommitDate")
        IniWrite(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss"), checkFile, "UpdateCheck", "LastCheckTime")

        pendingStr := ""
        if Type(pendingFiles) = "Array" {
            for item in pendingFiles {
                if IsExcludedPendingItem(item)
                    continue
                ; Legacy v5 updater-pending records are migrated into [UpdaterState]
                ; and must never be written back into generic PendingFiles.
                parsed := ParsePendingItem(item)
                if parsed.kind = "updater"
                    continue
                if GetUpdaterKindForRelPath(parsed.sourceRelPath) != ""
                    continue
                if GetUpdaterKindForRelPath(parsed.destRelPath) != ""
                    continue
                storedItem := NormalizePendingItemForStorage(item)
                pendingStr .= (pendingStr = "" ? "" : "|") storedItem
            }
        }
        IniWrite(pendingStr, checkFile, "UpdateCheck", "PendingFiles")

        updatedStr := ""
        if Type(updatedFiles) = "Array" {
            sortedUpdated := SortUpdatedFileRecords(updatedFiles)
            for record in sortedUpdated {
                if !record.HasOwnProp("path") or record.path = ""
                    continue
                if ShouldExcludeFromUpdates(record.path) or IsUpdaterSnapshotDestPath(record.path)
                    continue
                timestamp := record.HasOwnProp("timestamp") ? record.timestamp : A_Now
                action := record.HasOwnProp("action") and record.action != "" ? record.action : "Updated"
                updatedStr .= (updatedStr = "" ? "" : "|") timestamp "=>" action "=>" record.path
            }
        }
        IniWrite(updatedStr, checkFile, "UpdateCheck", "UpdatedFiles")

        LogDebug(
            "Stored update check info: SHA=" sha
            ", Date=" date
            ", Pending=" pendingStr
            ", Updated=" updatedStr)
    } catch Error as e {
        LogDebug("Error writing LastUpdateCheck.ini: " e.Message)
    }
}
ParseUpdatedFileRecords(raw) {
    ; Current format: timestamp=>action=>destinationPath.
    ; The v5 timestamp=>destinationPath format is accepted for migration.
    records := []
    if raw = ""
        return records

    for item in StrSplit(raw, "|") {
        item := Trim(item)
        if item = ""
            continue

        parts := StrSplit(item, "=>")
        if parts.Length < 2 {
            LogDebug("Ignoring malformed UpdatedFiles entry: " item)
            continue
        }

        timestamp := Trim(parts[1])
        if parts.Length = 2 {
            action := "Updated"
            path := Trim(parts[2])
        } else {
            action := Trim(parts[2])
            path := Trim(parts[3])
            ; Be defensive if a path ever contains the delimiter.
            if parts.Length > 3 {
                loop parts.Length - 3
                    path .= "=>" parts[A_Index + 3]
            }
        }

        if path = "" or ShouldExcludeFromUpdates(path) or IsUpdaterSnapshotDestPath(path)
            continue
        records.Push({path: path, timestamp: timestamp, action: action})
    }

    return SortUpdatedFileRecords(records)
}
UpsertUpdatedFileRecord(records, relPath, timestamp := "", action := "Updated") {
    ; Keep one completion record per generic destination path for this cycle.
    if Type(records) != "Array"
        records := []
    if ShouldExcludeFromUpdates(relPath) or IsUpdaterSnapshotDestPath(relPath)
        return records
    if timestamp = ""
        timestamp := A_Now

    indexToDelete := 0
    for index, record in records {
        if record.HasOwnProp("path") and StrLower(record.path) = StrLower(relPath) {
            indexToDelete := index
            break
        }
    }
    if indexToDelete
        records.RemoveAt(indexToDelete)

    records.Push({path: relPath, timestamp: timestamp, action: action})
    return records
}
SortUpdatedFileRecords(records) {
    ; Return a copy sorted newest-first. AHK timestamps are lexically sortable.
    sorted := []
    if Type(records) != "Array"
        return sorted

    for record in records {
        insertAt := sorted.Length + 1
        recordTime := record.HasOwnProp("timestamp") ? record.timestamp : ""

        for index, existing in sorted {
            existingTime := existing.HasOwnProp("timestamp") ? existing.timestamp : ""
            if recordTime > existingTime {
                insertAt := index
                break
            }
        }
        sorted.InsertAt(insertAt, record)
    }
    return sorted
}

FormatUpdatedTimestamp(timestamp) {
    ; Convert the stored sortable timestamp into a user-facing date/time.
    if RegExMatch(timestamp, "^\d{14}$") {
        try return FormatTime(timestamp, "dd MMM yyyy, HH:mm:ss")
    }
    return timestamp != "" ? timestamp : "Unknown"
}

MakePendingItem(sourceRelPath, destRelPath := "", changeType := "update") {
    ; Persist whether a generic pending file is a new file or an update so the
    ; pending GUI can preserve the same default-selection policy as the initial GUI.
    ; A source=>destination form is used when the downloaded repository path differs
    ; from the local target path.
    changeType := StrLower(changeType)
    if changeType != "new" and changeType != "update"
        changeType := "update"

    mapping := (
        destRelPath = "" or destRelPath = sourceRelPath
        ? sourceRelPath
        : sourceRelPath "=>" destRelPath)
    return "__PENDING_" StrUpper(changeType) "__" mapping
}

MakeHotstringReplacePendingItem(sourceRelPath, activeRelPath, snapshotRelPath) {
    ; Replace mode is a compound operation: upstream -> active library, then
    ; upstream -> configured upstream snapshot. Keep that intent across retries.
    return "__HOTSTRING_REPLACE__" sourceRelPath "=>" activeRelPath "=>" snapshotRelPath
}

ParseBasicPendingMapping(mapping) {
    separatorPos := InStr(mapping, "=>")
    if separatorPos {
        return {
            sourceRelPath: SubStr(mapping, 1, separatorPos - 1),
            destRelPath: SubStr(mapping, separatorPos + 2)
        }
    }
    return {sourceRelPath: mapping, destRelPath: mapping}
}

ParsePendingItem(pendingItem) {
    ; Decode current generic pending records plus legacy updater records.
    if RegExMatch(pendingItem, "^__PENDING_(NEW|UPDATE)__(.*)$", &match) {
        mapping := ParseBasicPendingMapping(match[2])
        return {
            kind: "normal",
            changeType: StrLower(match[1]),
            sourceRelPath: mapping.sourceRelPath,
            destRelPath: mapping.destRelPath
        }
    }

    if RegExMatch(pendingItem, "^__UPDATER_(AHK|EXE)__([0-9A-Fa-f]+)::(.*)$", &match) {
        mapping := ParseBasicPendingMapping(match[3])
        return {
            kind: "updater",
            updaterKind: StrLower(match[1]),
            upstreamHash: StrUpper(match[2]),
            sourceRelPath: mapping.sourceRelPath,
            destRelPath: mapping.destRelPath
        }
    }

    hotstringPrefix := "__HOTSTRING_REPLACE__"
    if SubStr(pendingItem, 1, StrLen(hotstringPrefix)) = hotstringPrefix {
        body := SubStr(pendingItem, StrLen(hotstringPrefix) + 1)
        parts := StrSplit(body, "=>")
        if parts.Length = 3 {
            return {
                kind: "hotstring_replace",
                changeType: "update",
                sourceRelPath: parts[1],
                destRelPath: parts[2],
                snapshotRelPath: parts[3]
            }
        }
    }

    ; Legacy generic pending entries from v6 and earlier did not persist whether
    ; the target was new or updated. Resolve that lazily from the current destination.
    mapping := ParseBasicPendingMapping(pendingItem)
    return {
        kind: "normal",
        changeType: "",
        sourceRelPath: mapping.sourceRelPath,
        destRelPath: mapping.destRelPath
    }
}

ResolvePendingChangeType(item) {
    ; Current v6.1 records carry the classification explicitly. For legacy records,
    ; infer it from whether the destination currently exists, then persist it the next
    ; time the state is written.
    if item.HasOwnProp("changeType") {
        changeType := StrLower(item.changeType)
        if changeType = "new" or changeType = "update"
            return changeType
    }

    SplitPath(A_ScriptDir, , &installDir)
    return FileExist(installDir "\" item.destRelPath) ? "update" : "new"
}

GetPendingChangeTypeLabel(item) {
    return ResolvePendingChangeType(item) = "new" ? "New" : "Update"
}

ShouldPendingItemStartChecked(item) {
    ; Match the initial GUI policy: new files are checked everywhere, while updates
    ; under rarely-updated folders (such as Data\) start unchecked.
    if ResolvePendingChangeType(item) = "new"
        return true
    return !IsRarelyUpdated(item.destRelPath, RarelyUpdatedFolders)
}

NormalizePendingItemForStorage(pendingItem) {
    item := ParsePendingItem(pendingItem)
    if item.kind != "normal"
        return pendingItem
    return MakePendingItem(
        item.sourceRelPath,
        item.destRelPath,
        ResolvePendingChangeType(item))
}

ShouldExcludeFromUpdates(relPath) {
    ; Runtime-generated state/log files should never be shown as update candidates.
    static excluded := Map()
    static initialized := false

    if !initialized {
        for path in ExcludedRuntimeFiles
            excluded[NormalizePath(path)] := true
        initialized := true
    }

    return excluded.Has(NormalizePath(relPath))
}

IsExcludedPendingItem(pendingItem) {
    item := ParsePendingItem(pendingItem)
    return (
        ShouldExcludeFromUpdates(item.sourceRelPath)
        or ShouldExcludeFromUpdates(item.destRelPath)
        or (item.kind = "hotstring_replace" and ShouldExcludeFromUpdates(item.snapshotRelPath))
    )
}

GetUpdaterKindForRelPath(relPath) {
    normalized := NormalizePath(relPath)
    if normalized = "tools\updater.ahk"
        return "ahk"
    if normalized = "tools\updater.exe"
        return "exe"
    return ""
}

GetUpdaterSourceRelPath(updaterKind) {
    return updaterKind = "ahk" ? "Tools\Updater.ahk" : "Tools\Updater.exe"
}

GetUpdaterUpstreamDestRelPath(updaterKind) {
    return updaterKind = "ahk"
        ? "Tools\Updater_Upstream\Updater.ahk"
        : "Tools\Updater_Upstream\Updater.exe"
}

IsUpdaterSnapshotDestPath(relPath) {
    normalized := NormalizePath(relPath)
    return (
        normalized = "tools\updater_upstream\updater.ahk"
        or normalized = "tools\updater_upstream\updater.exe"
    )
}

GetUpdaterState() {
    ; Read persistent handled hashes and current-cycle updater state. For migration,
    ; handled hashes fall back to v5's [UpstreamFileState] keys when necessary.
    checkFile := A_ScriptDir "\..\Data\LastUpdateCheck.ini"
    state := Map()

    for updaterKind in ["ahk", "exe"] {
        prefix := updaterKind = "ahk" ? "Ahk" : "Exe"
        legacyKey := updaterKind = "ahk" ? "UpdaterAhkHash" : "UpdaterExeHash"

        handledHash := ""
        candidateHash := ""
        status := "none"
        completedAt := ""
        try handledHash := IniRead(checkFile, "UpdaterState", prefix "HandledHash", "")
        try candidateHash := IniRead(checkFile, "UpdaterState", prefix "CandidateHash", "")
        try status := StrLower(IniRead(checkFile, "UpdaterState", prefix "Status", "none"))
        try completedAt := IniRead(checkFile, "UpdaterState", prefix "CompletedAt", "")

        if handledHash = "" {
            try handledHash := IniRead(checkFile, "UpstreamFileState", legacyKey, "")
        }

        if status != "none" and status != "pending" and status != "completed"
            status := "none"

        state[updaterKind] := Map(
            "handledHash", StrUpper(handledHash),
            "candidateHash", StrUpper(candidateHash),
            "status", status,
            "completedAt", completedAt)
    }

    return state
}

StoreUpdaterState(updaterState) {
    checkFile := A_ScriptDir "\..\Data\LastUpdateCheck.ini"
    checkDir := A_ScriptDir "\..\Data"
    if !DirExist(checkDir)
        DirCreate(checkDir)

    for updaterKind in ["ahk", "exe"] {
        prefix := updaterKind = "ahk" ? "Ahk" : "Exe"
        item := updaterState[updaterKind]
        IniWrite(item["handledHash"], checkFile, "UpdaterState", prefix "HandledHash")
        IniWrite(item["candidateHash"], checkFile, "UpdaterState", prefix "CandidateHash")
        IniWrite(item["status"], checkFile, "UpdaterState", prefix "Status")
        IniWrite(item["completedAt"], checkFile, "UpdaterState", prefix "CompletedAt")
    }

    LogDebug(
        "Stored UpdaterState: AhkStatus=" updaterState["ahk"]["status"]
        ", ExeStatus=" updaterState["exe"]["status"])
}

ResetUpdaterCycleState(updaterState) {
    ; A new global repository SHA starts a fresh updater-file cycle while preserving
    ; the last successfully handled upstream hashes.
    for updaterKind in ["ahk", "exe"] {
        item := updaterState[updaterKind]
        item["candidateHash"] := ""
        item["status"] := "none"
        item["completedAt"] := ""
    }
    return updaterState
}

SetUpdaterCandidate(updaterState, updaterKind, candidateHash) {
    item := updaterState[updaterKind]
    candidateHash := StrUpper(candidateHash)
    item["candidateHash"] := candidateHash
    item["completedAt"] := ""
    item["status"] := (
        item["handledHash"] != ""
        and StrUpper(item["handledHash"]) = candidateHash
        ? "none"
        : "pending")
}

MarkUpdaterCompleted(updaterState, updaterKind, timestamp := "") {
    if timestamp = ""
        timestamp := A_Now
    item := updaterState[updaterKind]
    item["handledHash"] := item["candidateHash"]
    item["status"] := "completed"
    item["completedAt"] := timestamp
}

IsUpdaterCycleStateInitialized(updaterState) {
    ; Validate not only field presence but the invariants implied by each status.
    for updaterKind in ["ahk", "exe"] {
        item := updaterState[updaterKind]
        candidate := StrUpper(item["candidateHash"])
        handled := StrUpper(item["handledHash"])
        status := item["status"]

        if candidate = ""
            return false

        if status = "none" {
            if handled = "" or handled != candidate
                return false
        } else if status = "pending" {
            if handled != "" and handled = candidate
                return false
        } else if status = "completed" {
            if handled = "" or handled != candidate or item["completedAt"] = ""
                return false
        } else {
            return false
        }
    }
    return true
}

UpdaterHasPending(updaterState) {
    return (
        updaterState["ahk"]["status"] = "pending"
        or updaterState["exe"]["status"] = "pending"
    )
}

GetUpdaterPendingCount(updaterState) {
    count := 0
    for updaterKind in ["ahk", "exe"] {
        if updaterState[updaterKind]["status"] = "pending"
            count += 1
    }
    return count
}

MigrateLegacyUpdaterPending(pendingFiles, updaterState) {
    ; v5 stored updater snapshots inside generic PendingFiles. Move those records
    ; into [UpdaterState] in memory and return only true generic pending records.
    genericPending := []
    if Type(pendingFiles) != "Array"
        return genericPending

    for pendingKey in pendingFiles {
        item := ParsePendingItem(pendingKey)
        if item.kind = "updater" {
            stateItem := updaterState[item.updaterKind]
            stateItem["candidateHash"] := StrUpper(item.upstreamHash)
            stateItem["status"] := "pending"
            stateItem["completedAt"] := ""
            LogDebug("Migrated legacy updater pending state: " item.sourceRelPath)
            continue
        }

        ; Older updater versions could leave the active updater paths as ordinary
        ; pending records. Never allow those records to overwrite customized files.
        updaterKind := GetUpdaterKindForRelPath(item.sourceRelPath)
        if updaterKind = ""
            updaterKind := GetUpdaterKindForRelPath(item.destRelPath)
        if updaterKind != "" {
            stateItem := updaterState[updaterKind]
            stateItem["candidateHash"] := ""
            stateItem["status"] := "pending"
            stateItem["completedAt"] := ""
            LogDebug("Removed legacy active-updater path from generic PendingFiles: " pendingKey)
            continue
        }

        genericPending.Push(pendingKey)
    }
    return genericPending
}

GetCheckedRowSet(lv) {
    ; Return a Map whose keys are the checked row numbers in a ListView.
    checked := Map()
    if lv = ""
        return checked

    rowNum := 0
    loop {
        rowNum := lv.GetNext(rowNum, "C")
        if rowNum = 0
            break
        checked[rowNum] := true
    }
    return checked
}

SetListViewChecks(lv, shouldCheck) {
    ; Check or uncheck every row in a selectable ListView.
    if lv = ""
        return

    loop lv.GetCount()
        lv.Modify(A_Index, shouldCheck ? "+Check" : "-Check")
}

SetMainUpdateChecks(updateGui, shouldCheck, *) {
    ; Updater snapshots are read-only/automatic and are intentionally not affected.
    SetListViewChecks(updateGui.lvUpdated, shouldCheck)
    SetListViewChecks(updateGui.lvNew, shouldCheck)
    SetListViewChecks(updateGui.lvHotstringLib, shouldCheck)
}
SetPendingChecks(pendingGui, shouldCheck, *) {
    ; Apply Check All / Uncheck All to the pending list only.
    SetListViewChecks(pendingGui.lvPending, shouldCheck)
}

CreateUpdaterGui(title, fontSize, formColor, fontColor, listColor) {
    ; Shared base builder for the initial-update and pending-update windows.
    updaterGui := Gui()
    updaterGui.Title := title
    updaterGui.Opt("+AlwaysOnTop")
    updaterGui.BackColor := formColor
    updaterGui.SetFont("s" fontSize " " fontColor)

    updaterGui.fontSize := fontSize
    updaterGui.formColor := formColor
    updaterGui.fontColor := fontColor
    updaterGui.listColor := listColor
    return updaterGui
}

AddSelectionButtons(updaterGui, checkHandler, uncheckHandler) {
    ; Shared Check All / Uncheck All controls used by both updater workflows.
    updaterGui.Add("Button", "w70 h24 y+6", "Check All").OnEvent("Click", checkHandler)
    updaterGui.Add("Button", "x+6 w82 h24", "Uncheck All").OnEvent("Click", uncheckHandler)
}

GetCompletedRecords(updatedFiles, updaterState) {
    ; Combine generic cycle completions with updater snapshot completions for display.
    records := []
    if Type(updatedFiles) = "Array" {
        for record in updatedFiles
            records.Push(record)
    }

    for updaterKind in ["ahk", "exe"] {
        item := updaterState[updaterKind]
        if item["status"] = "completed" and item["completedAt"] != "" {
            records.Push({
                path: GetUpdaterUpstreamDestRelPath(updaterKind),
                timestamp: item["completedAt"],
                action: "Upstream snapshot saved"
            })
        }
    }

    return SortUpdatedFileRecords(records)
}

AddCompletedFilesReadOnlySection(updaterGui, updatedFiles, updaterState) {
    ; Pending-mode-only completion summary. Do not create an empty section.
    completed := GetCompletedRecords(updatedFiles, updaterState)
    if completed.Length = 0
        return ""

    heading := updaterGui.Add(
        "Text",
        "xm w760 cGreen y+12",
        "Completed in this update (" completed.Length "):")
    updaterGui.completedHeading := heading

    lvCompleted := updaterGui.Add(
        "ListView",
        "w760 r6 -Multi +Background" updaterGui.listColor " y+5",
        ["Filename", "Path", "Action", "Completed at"])

    for record in completed {
        SplitPath(record.path, &fileName)
        action := record.HasOwnProp("action") ? record.action : "Updated"
        lvCompleted.Add("", fileName, record.path, action, FormatUpdatedTimestamp(record.timestamp))
    }

    lvCompleted.ModifyCol(1, 155)
    lvCompleted.ModifyCol(2, 255)
    lvCompleted.ModifyCol(3, 160)
    lvCompleted.ModifyCol(4, 190)
    updaterGui.lvCompleted := lvCompleted
    return lvCompleted
}

AddUpdaterSnapshotSection(updaterGui, updaterState, headingText) {
    ; Read-only and conditional. Updater snapshots are always processed automatically.
    pendingCount := GetUpdaterPendingCount(updaterState)
    if pendingCount = 0
        return ""

    updaterGui.Add("Text", "xm w760 cBlue y+10", headingText " (" pendingCount "):")
    updaterGui.Add(
        "Text",
        "xm w760",
        "These safe upstream copies are processed automatically and never overwrite your customized updater.")

    lv := updaterGui.Add(
        "ListView",
        "w760 r4 -Multi +Background" updaterGui.listColor " y+5",
        ["Upstream file", "Save as"])

    for updaterKind in ["ahk", "exe"] {
        if updaterState[updaterKind]["status"] = "pending" {
            lv.Add("", GetUpdaterSourceRelPath(updaterKind), GetUpdaterUpstreamDestRelPath(updaterKind))
        }
    }
    lv.ModifyCol(1, 300)
    lv.ModifyCol(2, 460)
    return lv
}
GetLastLineNumber(filePath) {
    ; Return the final line number in a text file, using 1 for an empty file.
    lineNum := 0
    Loop Read, filePath
        lineNum := A_Index
    return lineNum > 0 ? lineNum : 1
}

IsNotepadEditor(editorPath) {
    ; Determine whether DefaultEditor already names notepad.exe.
    if Trim(editorPath) = ""
        return false
    normalized := StrReplace(Trim(editorPath), Chr(34), "")
    SplitPath(normalized, &fileName)
    return StrLower(fileName) = "notepad.exe"
}

OpenDebugLog(*) {
    ; Open Updater_debug.txt with the configured editor, preferably at its last line.
    settingsPath := A_ScriptDir "\..\Data\acSettings.ini"
    logPath := A_ScriptDir "\..\Debug\Updater_debug.txt"

    if !FileExist(logPath) {
        AcMsgBox.Show("Updater debug log not found:`n" logPath, "AutoCorrect2 Updater")
        return
    }

    editorCmd := IniRead(settingsPath, "Shared", "EditorCmd", "")
    defaultEditor := IniRead(settingsPath, "Shared", "DefaultEditor", "")
    lastLine := GetLastLineNumber(logPath)
    errors := []

    if editorCmd != "" {
        if !InStr(editorCmd, "{file}") or !InStr(editorCmd, "{line}") {
            errors.Push("EditorCmd is missing the required {file} and/or {line} placeholder.")
        } else {
            cmd := StrReplace(editorCmd, "{file}", logPath)
            cmd := StrReplace(cmd, "{line}", lastLine)
            try {
                Run(cmd)
                return
            } catch Error as e {
                errors.Push("EditorCmd failed:`n" e.Message)
            }
        }
    }

    if defaultEditor != "" {
        q := Chr(34)
        try {
            Run(q defaultEditor q " " q logPath q)
            return
        } catch Error as e {
            errors.Push("DefaultEditor failed:`n" e.Message)
        }

        ; Avoid retrying Notepad when DefaultEditor already was Notepad.
        if IsNotepadEditor(defaultEditor) {
            ShowOpenDebugLogError(logPath, editorCmd, defaultEditor, errors)
            return
        }
    }

    q := Chr(34)
    try {
        Run("notepad.exe " q logPath q)
        return
    } catch Error as e {
        errors.Push("notepad.exe fallback failed:`n" e.Message)
    }

    ShowOpenDebugLogError(logPath, editorCmd, defaultEditor, errors)
}

ShowOpenDebugLogError(logPath, editorCmd, defaultEditor, errors) {
    ; Report all attempted editor-launch failures without hiding the useful details.
    msg := "Unable to open the updater debug log.`n`nDebug file:`n" logPath
    msg .= "`n`nConfigured EditorCmd:`n" (editorCmd != "" ? editorCmd : "<empty>")
    msg .= "`n`nConfigured DefaultEditor:`n" (defaultEditor != "" ? defaultEditor : "<empty>")

    if errors.Length > 0 {
        msg .= "`n`nAttempts:"
        for errorText in errors
            msg .= "`n`n" errorText
    }

    AcMsgBox.Show(msg, "AutoCorrect2 Updater - Open Debug Log")
}

BuildCopyFailureDetailsText(failures) {
    details := ""
    for index, failure in failures {
        if index > 1
            details .= "`r`n`r`n----------------------------------------`r`n`r`n"
        details .= "File: " failure.displayPath "`r`n"
        details .= "Source: " failure.source "`r`n"
        details .= "Destination: " failure.destination "`r`n`r`n"
        details .= "AutoHotkey error:`r`n" failure.reason

        if failure.HasOwnProp("windowsErrorCode") {
            details .= "`r`n`r`nWindows error code:`r`n" failure.windowsErrorCode
            details .= "`r`n`r`nWindows error:`r`n"
            details .= (
                failure.HasOwnProp("windowsErrorMessage") and failure.windowsErrorMessage != ""
                ? failure.windowsErrorMessage
                : "No additional Windows error message was available.")
        }
    }
    return details
}

ColorToBgr(colorValue) {
    color := Trim(colorValue)
    if SubStr(color, 1, 1) = "c"
        color := SubStr(color, 2)
    if SubStr(color, 1, 2) = "0x"
        color := SubStr(color, 3)
    if SubStr(color, 1, 1) = "#"
        color := SubStr(color, 2)
    color := Format("{:06s}", color)
    color := StrReplace(color, " ", "0")
    rgb := ("0x" color) + 0
    red := (rgb >> 16) & 0xFF
    green := (rgb >> 8) & 0xFF
    blue := rgb & 0xFF
    return (blue << 16) | (green << 8) | red
}

ApplyThemedEditColors(editCtrl, bgColor) {
    ; Best-effort themed background for the details Edit control.
    try {
        SendMessage(0x443, 0, ColorToBgr(bgColor), editCtrl) ; EM_SETBKGNDCOLOR
    }
}

ToggleCopyFailureDetails(failureGui, *) {
    if failureGui.detailsVisible {
        if failureGui.HasOwnProp("detailsEdit")
            failureGui.detailsEdit.Visible := false
        failureGui.detailsToggleButton.Text := "Show Details"
        failureGui.detailsVisible := false
        failureGui.Show("w" failureGui.compactWidth " h" failureGui.compactHeight)
        return
    }

    if !failureGui.HasOwnProp("detailsEdit") {
        detailsEdit := failureGui.Add(
            "Edit",
            "xm w760 r18 ReadOnly +VScroll " failureGui.detailsFontColor
            " Background" failureGui.detailsFormColor,
            failureGui.detailsText)
        ApplyThemedEditColors(detailsEdit, failureGui.detailsFormColor)
        failureGui.detailsEdit := detailsEdit
    } else {
        failureGui.detailsEdit.Visible := true
    }

    failureGui.detailsToggleButton.Text := "Hide Details"
    failureGui.detailsVisible := true
    failureGui.Show("AutoSize")
}

ShowCopyFailureDialog(failures, fontSize, formColor, fontColor, listColor) {
    ; Keep the updater open and present one compact failure window whose details can
    ; be expanded/collapsed in place.
    failureGui := Gui()
    failureGui.Title := "AutoCorrect2 Updater - Update Failed"
    failureGui.Opt("+AlwaysOnTop")
    failureGui.BackColor := formColor
    failureGui.SetFont("s" fontSize " " fontColor)

    message := "The following file(s) could not be processed:`n`n"
    for failure in failures
        message .= "  - " failure.displayPath "`n"
    message .= "`nThey remain pending and available for retry."

    failureGui.Add("Text", "w620", message)
    failureGui.Add("Button", "xm w100 y+10 Default", "OK").OnEvent("Click", (*) => failureGui.Destroy())
    toggleButton := failureGui.Add("Button", "x+10 w120", "Show Details")

    failureGui.detailsText := BuildCopyFailureDetailsText(failures)
    failureGui.detailsToggleButton := toggleButton
    failureGui.detailsVisible := false
    failureGui.detailsFormColor := formColor
    failureGui.detailsFontColor := fontColor
    toggleButton.OnEvent("Click", ToggleCopyFailureDetails.Bind(failureGui))

    failureGui.Show("AutoSize")
    failureGui.GetPos(, , &compactWidth, &compactHeight)
    failureGui.compactWidth := compactWidth
    failureGui.compactHeight := compactHeight
}

FilesHaveSameContent(firstPath, secondPath) {
    ; Fast path: different byte counts guarantee different contents.
    firstSize := FileGetSize(firstPath)
    secondSize := FileGetSize(secondPath)
    if firstSize != secondSize {
        LogDebug("File comparison: sizes differ (" firstSize " vs " secondSize ")")
        return false
    }

    ; Equal size does not prove equality, so compare SHA-256 hashes.
    firstHash := GetFileSha256(firstPath)
    secondHash := GetFileSha256(secondPath)
    hashesMatch := firstHash = secondHash

    LogDebug("File comparison: sizes match; SHA-256 hashes "
        (hashesMatch ? "match" : "differ"))
    return hashesMatch
}

GetFileSha256(filePath) {
    ; Compute a SHA-256 digest with the Windows CNG (bcrypt) API.
    hAlg := 0
    hHash := 0
    file := FileOpen(filePath, "r")
    if !file
        throw Error("Unable to open file for SHA-256 hashing: " filePath)

    try {
        status := DllCall(
            "bcrypt\BCryptOpenAlgorithmProvider",
            "Ptr*", &hAlg,
            "WStr", "SHA256",
            "Ptr", 0,
            "UInt", 0,
            "UInt"
        )
        if status != 0
            throw Error("BCryptOpenAlgorithmProvider failed with status " status)

        objectLengthBuffer := Buffer(4, 0)
        bytesWritten := 0
        status := DllCall(
            "bcrypt\BCryptGetProperty",
            "Ptr", hAlg,
            "WStr", "ObjectLength",
            "Ptr", objectLengthBuffer.Ptr,
            "UInt", objectLengthBuffer.Size,
            "UInt*", &bytesWritten,
            "UInt", 0,
            "UInt"
        )
        if status != 0
            throw Error("BCryptGetProperty(ObjectLength) failed with status " status)
        objectLength := NumGet(objectLengthBuffer, 0, "UInt")

        hashLengthBuffer := Buffer(4, 0)
        bytesWritten := 0
        status := DllCall(
            "bcrypt\BCryptGetProperty",
            "Ptr", hAlg,
            "WStr", "HashDigestLength",
            "Ptr", hashLengthBuffer.Ptr,
            "UInt", hashLengthBuffer.Size,
            "UInt*", &bytesWritten,
            "UInt", 0,
            "UInt"
        )
        if status != 0
            throw Error("BCryptGetProperty(HashDigestLength) failed with status " status)
        hashLength := NumGet(hashLengthBuffer, 0, "UInt")

        hashObject := Buffer(objectLength, 0)
        status := DllCall(
            "bcrypt\BCryptCreateHash",
            "Ptr", hAlg,
            "Ptr*", &hHash,
            "Ptr", hashObject.Ptr,
            "UInt", hashObject.Size,
            "Ptr", 0,
            "UInt", 0,
            "UInt", 0,
            "UInt"
        )
        if status != 0
            throw Error("BCryptCreateHash failed with status " status)

        chunk := Buffer(1024 * 1024, 0)
        loop {
            bytesRead := file.RawRead(chunk, chunk.Size)
            if bytesRead = 0
                break

            status := DllCall(
                "bcrypt\BCryptHashData",
                "Ptr", hHash,
                "Ptr", chunk.Ptr,
                "UInt", bytesRead,
                "UInt", 0,
                "UInt"
            )
            if status != 0
                throw Error("BCryptHashData failed with status " status)
        }

        hashBuffer := Buffer(hashLength, 0)
        status := DllCall(
            "bcrypt\BCryptFinishHash",
            "Ptr", hHash,
            "Ptr", hashBuffer.Ptr,
            "UInt", hashBuffer.Size,
            "UInt", 0,
            "UInt"
        )
        if status != 0
            throw Error("BCryptFinishHash failed with status " status)

        digest := ""
        loop hashBuffer.Size
            digest .= Format("{:02X}", NumGet(hashBuffer, A_Index - 1, "UChar"))
        return digest
    } finally {
        if hHash
            DllCall("bcrypt\BCryptDestroyHash", "Ptr", hHash)
        if hAlg
            DllCall("bcrypt\BCryptCloseAlgorithmProvider", "Ptr", hAlg, "UInt", 0)
        file.Close()
    }
}

GetShaCacheDir(sha) {
    return A_Temp "\AutoCorrect2_Update\" sha
}

FindCachedSourceRoot(extractDir) {
    if !DirExist(extractDir)
        return ""

    loop Files, extractDir "\*", "D" {
        candidate := A_LoopFileFullPath
        if DirExist(candidate "\Core") and DirExist(candidate "\Tools")
            return candidate
    }
    return ""
}

GetShaCacheMarkerPath(sha) {
    return GetShaCacheDir(sha) "\complete.marker"
}

GetSourceTreeSignature(srcRoot) {
    ; Count/size signature used only while waiting for Shell.Application extraction
    ; to stop changing. The completion marker is what makes a cache reusable later.
    if srcRoot = "" or !DirExist(srcRoot)
        return ""
    fileCount := 0
    totalSize := 0
    loop Files, srcRoot "\*", "FR" {
        fileCount += 1
        totalSize += A_LoopFileSize
    }
    return fileCount > 0 ? fileCount ":" totalSize : ""
}

WaitForExtractionToStabilize(extractDir) {
    lastSignature := ""
    stableChecks := 0
    loop 300 {
        srcRoot := FindCachedSourceRoot(extractDir)
        if srcRoot != "" {
            signature := GetSourceTreeSignature(srcRoot)
            if signature != "" and signature = lastSignature {
                stableChecks += 1
                if stableChecks >= 20
                    return srcRoot
            } else {
                stableChecks := 0
                lastSignature := signature
            }
        }
        Sleep(100)
    }
    return ""
}

SourceHasRequiredFiles(srcRoot, requiredRelPaths := "") {
    if srcRoot = "" or !DirExist(srcRoot)
        return false
    if Type(requiredRelPaths) != "Array"
        return true

    for relPath in requiredRelPaths {
        if relPath = ""
            continue
        if !FileExist(srcRoot "\" relPath)
            return false
    }
    return true
}

InvalidateShaCache(sha) {
    cacheDir := GetShaCacheDir(sha)
    if DirExist(cacheDir) {
        LogDebug("Invalidating SHA cache: " cacheDir)
        DirDelete(cacheDir, 1)
    }
}

EnsureSourceForSha(sha, requiredRelPaths := "", progressGui := "", statusTextCtrl := "", progressBarCtrl := "") {
    ; Reuse an exact-SHA cache when complete. Otherwise rebuild and download that
    ; precise commit once, never the moving main branch.
    cacheDir := GetShaCacheDir(sha)
    zipFile := cacheDir "\archive.zip"
    extractDir := cacheDir "\extracted"
    srcRoot := FindCachedSourceRoot(extractDir)
    markerPath := GetShaCacheMarkerPath(sha)

    if FileExist(markerPath) and SourceHasRequiredFiles(srcRoot, requiredRelPaths) {
        LogDebug("Reusing exact-SHA source cache: " sha)
        return srcRoot
    }

    if DirExist(cacheDir) {
        LogDebug("SHA cache missing/incomplete required source files; rebuilding: " sha)
        DirDelete(cacheDir, 1)
    }
    DirCreate(cacheDir)
    DirCreate(extractDir)

    archiveUrl := RepositoryArchiveBaseUrl "/" sha ".zip"
    LogDebug("Downloading exact commit archive: " archiveUrl)
    if progressGui != ""
        UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Downloading exact GitHub revision...", 60)

    try Download(archiveUrl, zipFile)
    catch Error as e {
        throw Error("Failed to download exact GitHub revision " sha ".`n`nError: " e.Message)
    }

    if progressGui != ""
        UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Extracting exact GitHub revision...", 80)

    sh := ComObject("Shell.Application")
    zip := sh.NameSpace(zipFile)
    if !zip
        throw Error("Failed to open downloaded zip file for commit " sha ".")
    dest := sh.NameSpace(extractDir)
    if !dest
        throw Error("Failed to open extraction folder for commit " sha ".")
    dest.CopyHere(zip.Items, 0x10|0x4)

    ; Shell extraction is asynchronous. Wait until the extracted tree stops
    ; changing, then validate every required source file and write a completion marker.
    srcRoot := WaitForExtractionToStabilize(extractDir)
    if srcRoot = ""
        throw Error("Timed out waiting for extraction of commit " sha " to finish.")
    if !SourceHasRequiredFiles(srcRoot, requiredRelPaths)
        throw Error("The exact-SHA archive is missing one or more required source files for commit " sha ".")

    FileAppend(sha, markerPath)
    LogDebug("Exact-SHA source cache ready: " sha " -> " srcRoot)
    return srcRoot
}

CollectPendingSourcePaths(pendingFiles, updaterState) {
    required := []
    seen := Map()

    if Type(pendingFiles) = "Array" {
        for pendingKey in pendingFiles {
            item := ParsePendingItem(pendingKey)
            relPath := item.sourceRelPath
            key := NormalizePath(relPath)
            if relPath != "" and !seen.Has(key) {
                seen[key] := true
                required.Push(relPath)
            }
        }
    }

    for updaterKind in ["ahk", "exe"] {
        if updaterState[updaterKind]["status"] = "pending" {
            relPath := GetUpdaterSourceRelPath(updaterKind)
            key := NormalizePath(relPath)
            if !seen.Has(key) {
                seen[key] := true
                required.Push(relPath)
            }
        }
    }
    return required
}

PendingUpdaterHashesMatch(updaterState, srcRoot) {
    for updaterKind in ["ahk", "exe"] {
        item := updaterState[updaterKind]
        if item["status"] != "pending"
            continue

        sourcePath := srcRoot "\" GetUpdaterSourceRelPath(updaterKind)
        if !FileExist(sourcePath)
            return false
        if item["candidateHash"] = ""
            return false
        if StrUpper(GetFileSha256(sourcePath)) != StrUpper(item["candidateHash"])
            return false
    }
    return true
}

EnsurePendingSourceForCycle(sha, pendingFiles, updaterState, progressGui := "", statusTextCtrl := "", progressBarCtrl := "") {
    ; Validate the exact-SHA cache against every generic pending source plus every
    ; pending updater source. Updater candidate hashes provide an extra integrity check.
    required := CollectPendingSourcePaths(pendingFiles, updaterState)
    srcRoot := EnsureSourceForSha(sha, required, progressGui, statusTextCtrl, progressBarCtrl)

    if PendingUpdaterHashesMatch(updaterState, srcRoot)
        return srcRoot

    LogDebug("Pending updater candidate hash mismatch; rebuilding exact-SHA cache once.")
    InvalidateShaCache(sha)
    srcRoot := EnsureSourceForSha(sha, required, progressGui, statusTextCtrl, progressBarCtrl)
    if !PendingUpdaterHashesMatch(updaterState, srcRoot)
        throw Error("The updater source files for commit " sha " do not match the candidate hashes stored for this update cycle.")
    return srcRoot
}

InitializeUpdaterStateFromSource(updaterState, srcRoot) {
    ; Initialize/repair current-cycle updater state for an already-known repository SHA.
    for updaterKind in ["ahk", "exe"] {
        sourcePath := srcRoot "\" GetUpdaterSourceRelPath(updaterKind)
        if !FileExist(sourcePath)
            throw Error("Required updater source file is missing: " sourcePath)
        SetUpdaterCandidate(updaterState, updaterKind, GetFileSha256(sourcePath))
    }
    return updaterState
}

BuildUpdaterSpecialUpdates(updaterState, srcRoot, installDir) {
    items := []
    for updaterKind in ["ahk", "exe"] {
        stateItem := updaterState[updaterKind]
        if stateItem["status"] != "pending"
            continue
        sourceRelPath := GetUpdaterSourceRelPath(updaterKind)
        destRelPath := GetUpdaterUpstreamDestRelPath(updaterKind)
        items.Push({
            updaterKind: updaterKind,
            upstreamHash: stateItem["candidateHash"],
            sourceRelPath: sourceRelPath,
            destRelPath: destRelPath,
            path: installDir "\" destRelPath,
            srcPath: srcRoot "\" sourceRelPath
        })
    }
    return items
}

ScanRepositoryForUpdates(srcRoot, installDir, updaterState, hotstringLibName) {
    updatedFiles := []
    rarelyUpdatedFiles := []
    newFiles := []
    hotstringLibUpdate := ""

    loop Files, srcRoot "\*", "FR" {
        srcFile := A_LoopFileFullPath
        if InStr(A_LoopFileAttrib, "D")
            continue

        relPath := SubStr(srcFile, StrLen(srcRoot) + 2)
        destPath := installDir "\" relPath

        if ShouldExcludeFromUpdates(relPath) {
            LogDebug("Skipping excluded runtime/state file: " relPath)
            continue
        }

        updaterKind := GetUpdaterKindForRelPath(relPath)
        if updaterKind != "" {
            SetUpdaterCandidate(updaterState, updaterKind, GetFileSha256(srcFile))
            LogDebug(
                "Updater candidate " relPath ": status=" updaterState[updaterKind]["status"])
            continue
        }

        if relPath = "Core\AutoCorrectHotstrings.ahk" {
            snapshotPath := installDir "\Core\" hotstringLibName
            snapshotExists := FileExist(snapshotPath)
            if !snapshotExists or !FilesHaveSameContent(srcFile, snapshotPath) {
                hotstringLibUpdate := {
                    path: snapshotPath,
                    relPath: "Core\" hotstringLibName,
                    srcPath: srcFile,
                    isNew: !snapshotExists
                }
                LogDebug("Hotstring upstream snapshot needs refresh: " hotstringLibName)
            } else {
                LogDebug("Remote AutoCorrectHotstrings.ahk matches local upstream snapshot")
            }
            continue
        }

        if !FileExist(destPath) {
            newFiles.Push({path: destPath, relPath: relPath, srcPath: srcFile})
            LogDebug("New file found: " relPath)
            continue
        }

        if !FilesHaveSameContent(srcFile, destPath) {
            if IsRarelyUpdated(relPath, RarelyUpdatedFolders) {
                rarelyUpdatedFiles.Push({path: destPath, relPath: relPath, srcPath: srcFile})
                LogDebug("Rarely-updated file found: " relPath)
            } else {
                updatedFiles.Push({path: destPath, relPath: relPath, srcPath: srcFile})
                LogDebug("Updated file found: " relPath)
            }
        }
    }

    return Map(
        "updatedFiles", updatedFiles,
        "rarelyUpdatedFiles", rarelyUpdatedFiles,
        "newFiles", newFiles,
        "hotstringLibUpdate", hotstringLibUpdate,
        "updaterSpecialUpdates", BuildUpdaterSpecialUpdates(updaterState, srcRoot, installDir))
}

UpdateProgress(gui, textCtrl, progressBar, statusText, progressValue := 0) {
    gui.Title := "AutoCorrect2 Updater - " statusText
    textCtrl.Value := statusText
    progressBar.Value := progressValue
}

try {
    LogDebug("=== Update Started ===")
    LogDebug("Script Directory: " A_ScriptDir)

    settingsFile := A_ScriptDir "\..\Data\acSettings.ini"
    fontSize := IniRead(settingsFile, "HotstringHelper", "DefaultFontSize", "11")
    largeFontSize := IniRead(settingsFile, "HotstringHelper", "LargeFontSize", "13")
    hotstringLibName := IniRead(settingsFile, "Files", "NewTemporaryHotstrLib", "AutoCorrectHotstrings_Upstream.ahk")

    formColor := "0xE5E4E2"
    fontColor := "c0x1F1F1F"
    listColor := "0xFFFFFF"
    if FileExist(A_ScriptDir "\..\Data\colorThemeSettings.ini") {
        ctSettingsFile := A_ScriptDir "\..\Data\colorThemeSettings.ini"
        formColor := IniRead(ctSettingsFile, "ColorSettings", "formColor", formColor)
        fontColor := "c" IniRead(ctSettingsFile, "ColorSettings", "fontColor", "0x1F1F1F")
        listColor := IniRead(ctSettingsFile, "ColorSettings", "listColor", listColor)
    }

    class Config {
        static FormColor := formColor
        static FontColor := fontColor
        static LargeFontSize := "s" . largeFontSize
    }

    progressGui := Gui()
    progressGui.Title := "AutoCorrect2 Updater"
    progressGui.Opt("+AlwaysOnTop")
    progressGui.BackColor := formColor
    progressGui.SetFont("s" fontSize " " fontColor)
    statusTextCtrl := progressGui.Add("Text", "w400 h30 cBlue", "Checking for updates...")
    progressBarCtrl := progressGui.Add("Progress", "w400 h20 vMyProgress Range0-100", 0)
    progressGui.Show("w420 h100")
    UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Checking GitHub for updates...", 10)

    latestInfo := CheckGitHubForUpdates()
    if Type(latestInfo) = "String" {
        progressGui.Destroy()
        AcMsgBox.Show(
            "Unable to check GitHub for updates:`n" latestInfo
            "`n`nPlease check your internet connection and try again.",
            "AutoCorrect2 Updater")
        ExitApp
    }

    installDir := parentDir
    lastCheckInfo := GetLastUpdateCheckInfo()
    updaterState := GetUpdaterState()
    lastSha := lastCheckInfo.Has("sha") ? lastCheckInfo["sha"] : ""
    pendingFromLast := lastCheckInfo.Has("pendingFiles") ? lastCheckInfo["pendingFiles"] : []
    updatedFromLast := lastCheckInfo.Has("updatedFiles") ? lastCheckInfo["updatedFiles"] : []

    ; Migrate any v5 updater-pending entries out of generic PendingFiles.
    migratedPending := MigrateLegacyUpdaterPending(pendingFromLast, updaterState)
    if migratedPending.Length != pendingFromLast.Length {
        pendingFromLast := migratedPending
        if lastSha != "" {
            lastCheckInfo["pendingFiles"] := pendingFromLast
            StoreUpdateCheckInfo(lastSha, lastCheckInfo["date"], pendingFromLast, updatedFromLast)
            StoreUpdaterState(updaterState)
        }
    }

    if lastSha != "" and lastSha = latestInfo["sha"] {
        ; Same repository revision: repair/initialize updater-cycle state if this is
        ; the first v6 run over state produced by an older updater.
        if !IsUpdaterCycleStateInitialized(updaterState) {
            LogDebug("Same SHA but UpdaterState is incomplete - initializing from exact SHA.")
            requiredUpdaterFiles := [GetUpdaterSourceRelPath("ahk"), GetUpdaterSourceRelPath("exe")]
            srcRoot := EnsureSourceForSha(
                lastSha,
                requiredUpdaterFiles,
                progressGui,
                statusTextCtrl,
                progressBarCtrl)
            InitializeUpdaterStateFromSource(updaterState, srcRoot)
            StoreUpdaterState(updaterState)
        }

        if pendingFromLast.Length > 0 or UpdaterHasPending(updaterState) {
            LogDebug(
                "Same SHA has unfinished work: normal pending=" pendingFromLast.Length
                ", updater pending=" GetUpdaterPendingCount(updaterState))
            UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Preparing pending update sources...", 35)
            srcRoot := EnsurePendingSourceForCycle(
                lastSha,
                pendingFromLast,
                updaterState,
                progressGui,
                statusTextCtrl,
                progressBarCtrl)
            progressGui.Destroy()
            ShowPendingGui(
                pendingFromLast,
                lastCheckInfo,
                updaterState,
                srcRoot,
                fontSize,
                formColor,
                fontColor,
                listColor)
        } else {
            progressGui.Destroy()
            LogDebug("No updates available - same SHA and no unfinished work.")
            AcMsgBox.Show("No updates available. You have the latest version!", "AutoCorrect2 Updater")
            ExitApp
        }
    } else {
        ; Missing/different SHA means a genuinely new cycle. Generic cycle state and
        ; updater candidate/status/completion state reset; handled updater hashes persist.
        LogDebug("Starting fresh update cycle for commit " latestInfo["sha"])
        ResetUpdaterCycleState(updaterState)
        UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Preparing exact GitHub revision...", 35)
        srcRoot := EnsureSourceForSha(
            latestInfo["sha"],
            "",
            progressGui,
            statusTextCtrl,
            progressBarCtrl)
        UpdateProgress(progressGui, statusTextCtrl, progressBarCtrl, "Analyzing changes...", 100)

        scan := ScanRepositoryForUpdates(srcRoot, installDir, updaterState, hotstringLibName)
        if !IsUpdaterCycleStateInitialized(updaterState)
            throw Error("The exact repository revision does not contain a complete/valid Updater.ahk and Updater.exe state.")
        updatedFiles := scan["updatedFiles"]
        rarelyUpdatedFiles := scan["rarelyUpdatedFiles"]
        newFiles := scan["newFiles"]
        hotstringLibUpdate := scan["hotstringLibUpdate"]
        updaterSpecialUpdates := scan["updaterSpecialUpdates"]

        hasWork := (
            updatedFiles.Length > 0
            or rarelyUpdatedFiles.Length > 0
            or newFiles.Length > 0
            or hotstringLibUpdate != ""
            or UpdaterHasPending(updaterState))

        if !hasWork {
            StoreUpdateCheckInfo(latestInfo["sha"], latestInfo["date"], [], [])
            StoreUpdaterState(updaterState)
            progressGui.Destroy()
            LogDebug("Fresh exact-SHA comparison found no work.")
            AcMsgBox.Show("No updates available.", "AutoCorrect2 Updater")
            ExitApp
        }

        progressGui.Destroy()
        ShowUpdateGui(
            updatedFiles,
            rarelyUpdatedFiles,
            newFiles,
            updaterSpecialUpdates,
            updaterState,
            hotstringLibUpdate,
            hotstringLibName,
            installDir,
            srcRoot,
            latestInfo,
            fontSize,
            formColor,
            fontColor,
            listColor)
    }
} catch Error as e {
    LogDebug("ERROR: " e.Message)
    try progressGui.Destroy()

    errorMsg := "Update failed"
    if InStr(e.Message, "internet connection") or InStr(e.Message, "Failed to download") {
        errorMsg := "Unable to connect to GitHub. Please check your internet connection and try again.`n`n" e.Message
    } else {
        errorMsg := e.Message
    }

    AcMsgBox.Show(errorMsg, "AutoCorrect2 Updater")
    ExitApp
}

ShowPendingGui(pendingFiles, lastCheckInfo, updaterState, srcRoot, fontSize, formColor, fontColor, listColor, checkedPendingKeys := "") {
    ; Same-SHA unfinished-work GUI. Generic pending files are selectable; pending
    ; updater snapshots are read-only and automatically retried by Update.
    pendingGui := CreateUpdaterGui(
        "AutoCorrect2 Updater - Pending Files",
        fontSize,
        formColor,
        fontColor,
        listColor)

    updatedFiles := (
        Type(lastCheckInfo) = "Map" and lastCheckInfo.Has("updatedFiles")
        ? lastCheckInfo["updatedFiles"]
        : [])

    lvPending := ""
    if pendingFiles.Length > 0 {
        pendingGui.Add(
            "Text",
            "w760 cBlue",
            "Pending files (" pendingFiles.Length "):")
        pendingGui.Add("Text", "w760", "Check the files you want to update now, then click Update.")

        lvPending := pendingGui.Add(
            "ListView",
            "w760 r10 Checked -Multi +Background" listColor " y+5",
            ["Filename", "Path", "Type", "PendingKey"])
        for pendingKey in pendingFiles {
            item := ParsePendingItem(pendingKey)
            SplitPath(item.destRelPath, &fileName)
            shouldCheck := (
                Type(checkedPendingKeys) = "Map"
                ? checkedPendingKeys.Has(pendingKey)
                : ShouldPendingItemStartChecked(item))
            lvPending.Add(
                shouldCheck ? "Check" : "",
                fileName,
                item.destRelPath,
                GetPendingChangeTypeLabel(item),
                pendingKey)
        }
        lvPending.ModifyCol(1, 180)
        lvPending.ModifyCol(2, 450)
        lvPending.ModifyCol(3, 110)
        lvPending.ModifyCol(4, 0)

        AddSelectionButtons(
            pendingGui,
            SetPendingChecks.Bind(pendingGui, true),
            SetPendingChecks.Bind(pendingGui, false))
    }
    pendingGui.lvPending := lvPending

    ; Conditional read-only special section; omitted entirely when nothing is pending.
    AddUpdaterSnapshotSection(pendingGui, updaterState, "Updater snapshots still pending")

    ; Conditional read-only completion section; omitted entirely when empty.
    AddCompletedFilesReadOnlySection(pendingGui, updatedFiles, updaterState)

    pendingGui.Add("Button", "xm w100 y+18", "Update").OnEvent(
        "Click",
        ApplyPendingUpdate.Bind(lvPending, lastCheckInfo, updaterState, srcRoot, pendingGui))
    pendingGui.Add("Button", "x+10 w120", "Open Debug Log").OnEvent("Click", OpenDebugLog)
    if pendingFiles.Length > 0 {
        pendingGui.Add("Button", "x+10 w130", "Clear && Dismiss").OnEvent(
            "Click",
            ClearPendingAndExit.Bind(lvPending, lastCheckInfo, updaterState, pendingGui))
    }
    pendingGui.Add("Button", "x+10 w100", "Cancel").OnEvent("Click", (*) => ExitApp())

    pendingGui.Show()
    return pendingGui
}
RefreshUpdatedFilesReadOnlySection(pendingGui, updatedFiles) {
    ; Retained for compatibility with older call sites. v6 rebuilds the pending GUI
    ; after partial failures so conditional read-only sections remain accurate.
    return
}
RefreshPendingSelectableList(lv, pendingFiles, checkedPendingKeys := "") {
    ; Rebuild the pending ListView while preserving intended checked state.
    lv.Delete()
    for pendingKey in pendingFiles {
        pendingItem := ParsePendingItem(pendingKey)
        SplitPath(pendingItem.destRelPath, &fileName)
        shouldCheck := (
            Type(checkedPendingKeys) = "Map"
            ? checkedPendingKeys.Has(pendingKey)
            : ShouldPendingItemStartChecked(pendingItem))
        lv.Add(
            shouldCheck ? "Check" : "",
            fileName,
            pendingItem.destRelPath,
            GetPendingChangeTypeLabel(pendingItem),
            pendingKey)
    }
}

ApplyPendingUpdate(lv, lastCheckInfo, updaterState, srcRoot, pendingGui, *) {
    ; Retry selected generic pending files plus every read-only pending updater snapshot.
    SplitPath(A_ScriptDir, , &parentDir)
    installDir := parentDir
    checkedSet := GetCheckedRowSet(lv)
    failures := []
    replacementPending := Map()
    successfulRows := Map()
    retryCheckedPending := Map()
    updatedFiles := (
        Type(lastCheckInfo) = "Map" and lastCheckInfo.Has("updatedFiles")
        ? lastCheckInfo["updatedFiles"]
        : [])

    if lv != "" {
        loop lv.GetCount() {
            rowNum := A_Index
            if !checkedSet.Has(rowNum)
                continue

            pendingKey := lv.GetText(rowNum, 4)
            item := ParsePendingItem(pendingKey)
            srcFile := srcRoot "\" item.sourceRelPath

            if !FileExist(srcFile) {
                failures.Push({
                    displayPath: item.destRelPath,
                    source: srcFile,
                    destination: installDir "\" item.destRelPath,
                    reason: "Source file is not available in the exact-SHA source cache."
                })
                continue
            }

            if item.kind = "hotstring_replace" {
                activeDest := installDir "\" item.destRelPath
                snapshotDest := installDir "\" item.snapshotRelPath

                activeResult := CopyFileWithDirCreation(srcFile, activeDest)
                if activeResult.skipped {
                    activeResult.displayPath := item.destRelPath
                    failures.Push(activeResult)
                    continue
                }

                UpsertUpdatedFileRecord(updatedFiles, item.destRelPath, A_Now, "Replaced")
                snapshotExistedBefore := FileExist(snapshotDest)
                snapshotResult := CopyFileWithDirCreation(srcFile, snapshotDest)
                if snapshotResult.skipped {
                    snapshotResult.displayPath := item.snapshotRelPath
                    failures.Push(snapshotResult)
                    replacementPending[rowNum] := MakePendingItem(
                        item.sourceRelPath,
                        item.snapshotRelPath,
                        snapshotExistedBefore ? "update" : "new")
                    continue
                }

                UpsertUpdatedFileRecord(updatedFiles, item.snapshotRelPath, A_Now, "Upstream snapshot saved")
                successfulRows[rowNum] := true
                continue
            }

            destPath := installDir "\" item.destRelPath
            result := CopyFileWithDirCreation(srcFile, destPath)
            if result.skipped {
                result.displayPath := item.destRelPath
                failures.Push(result)
                continue
            }

            changeType := ResolvePendingChangeType(item)
            action := (
                NormalizePath(item.sourceRelPath) = "core\autocorrecthotstrings.ahk"
                and NormalizePath(item.destRelPath) != "core\autocorrecthotstrings.ahk"
                ? "Upstream snapshot saved"
                : (changeType = "new" ? "Added" : "Updated"))
            UpsertUpdatedFileRecord(updatedFiles, item.destRelPath, A_Now, action)
            successfulRows[rowNum] := true
        }
    }

    ; Updater snapshots are mandatory safe copies and are retried automatically.
    for updaterKind in ["ahk", "exe"] {
        stateItem := updaterState[updaterKind]
        if stateItem["status"] != "pending"
            continue

        sourceRelPath := GetUpdaterSourceRelPath(updaterKind)
        destRelPath := GetUpdaterUpstreamDestRelPath(updaterKind)
        srcFile := srcRoot "\" sourceRelPath
        destPath := installDir "\" destRelPath
        result := CopyFileWithDirCreation(srcFile, destPath)
        if result.skipped {
            result.displayPath := destRelPath
            failures.Push(result)
        } else {
            MarkUpdaterCompleted(updaterState, updaterKind, A_Now)
            LogDebug("Saved upstream updater snapshot: " destRelPath)
        }
    }

    remainingPending := []
    if lv != "" {
        loop lv.GetCount() {
            rowNum := A_Index
            if successfulRows.Has(rowNum)
                continue
            originalKey := lv.GetText(rowNum, 4)
            pendingKey := replacementPending.Has(rowNum) ? replacementPending[rowNum] : originalKey
            if pendingKey = "" or IsExcludedPendingItem(pendingKey)
                continue
            remainingPending.Push(pendingKey)
            if checkedSet.Has(rowNum)
                retryCheckedPending[pendingKey] := true
        }
    }

    StoreUpdateCheckInfo(lastCheckInfo["sha"], lastCheckInfo["date"], remainingPending, updatedFiles)
    StoreUpdaterState(updaterState)

    if failures.Length > 0 {
        ; Rebuild the same pending workflow so conditional read-only sections reflect
        ; updater snapshots that may have just succeeded or remained pending.
        newState := Map(
            "sha", lastCheckInfo["sha"],
            "date", lastCheckInfo["date"],
            "pendingFiles", remainingPending,
            "updatedFiles", updatedFiles)
        guiFontSize := pendingGui.fontSize
        guiFormColor := pendingGui.formColor
        guiFontColor := pendingGui.fontColor
        guiListColor := pendingGui.listColor
        pendingGui.Destroy()
        ShowPendingGui(
            remainingPending,
            newState,
            updaterState,
            srcRoot,
            guiFontSize,
            guiFormColor,
            guiFontColor,
            guiListColor,
            retryCheckedPending)
        ShowCopyFailureDialog(failures, guiFontSize, guiFormColor, guiFontColor, guiListColor)
        return
    }

    pendingGui.Hide()
    unfinishedCount := remainingPending.Length + GetUpdaterPendingCount(updaterState)
    if unfinishedCount > 0 {
        AcMsgBox.Show(
            "Selected files updated successfully.`n`n"
            unfinishedCount " pending item(s) remain for later review.",
            "AutoCorrect2 Updater")
    } else {
        AcMsgBox.Show("All pending work has been completed successfully!", "AutoCorrect2 Updater")
    }
    ExitApp
}
ClearPendingAndExit(lvPending, lastCheckInfo, updaterState, pendingGui, *) {
    ; Clear only checked generic pending files. Read-only updater snapshots are not
    ; dismissible here and remain pending until their safe comparison copy succeeds.
    checkedSet := GetCheckedRowSet(lvPending)
    if checkedSet.Count = 0 {
        AcMsgBox.Show("No pending files are selected.", "AutoCorrect2 Updater")
        return
    }

    updatedFiles := lastCheckInfo.Has("updatedFiles") ? lastCheckInfo["updatedFiles"] : []
    remainingPending := []
    dismissedCount := 0

    loop lvPending.GetCount() {
        pendingKey := lvPending.GetText(A_Index, 4)
        if pendingKey = ""
            continue
        if checkedSet.Has(A_Index) {
            dismissedCount += 1
            continue
        }
        if !IsExcludedPendingItem(pendingKey)
            remainingPending.Push(pendingKey)
    }

    StoreUpdateCheckInfo(lastCheckInfo["sha"], lastCheckInfo["date"], remainingPending, updatedFiles)
    StoreUpdaterState(updaterState)

    pendingGui.Hide()
    unfinishedCount := remainingPending.Length + GetUpdaterPendingCount(updaterState)
    if unfinishedCount > 0 {
        AcMsgBox.Show(
            "Dismissed " dismissedCount " selected pending file(s).`n`n"
            unfinishedCount " pending item(s) remain for later review.",
            "AutoCorrect2 Updater")
    } else {
        AcMsgBox.Show(
            "Dismissed " dismissedCount " selected pending file(s).`n`nNo pending work remains.",
            "AutoCorrect2 Updater")
    }
    ExitApp
}
ShowUpdateGui(updatedFiles, rarelyUpdatedFiles, newFiles, updaterSpecialUpdates, updaterState, hotstringLibUpdate, hotstringLibName, installDir, srcRoot, latestInfo, fontSize, formColor, fontColor, listColor) {
    ; Fresh-cycle GUI. Updater snapshots are conditional read-only automatic work.
    updateGui := CreateUpdaterGui(
        "AutoCorrect2 - Select Updates",
        fontSize,
        formColor,
        fontColor,
        listColor)

    lvUpdated := ""
    lvNew := ""
    lvHotstringLib := ""

    allUpdatedFiles := []
    for item in updatedFiles
        allUpdatedFiles.Push({data: item, shouldCheck: true})
    for item in rarelyUpdatedFiles
        allUpdatedFiles.Push({data: item, shouldCheck: false})

    if allUpdatedFiles.Length > 0 {
        updateGui.Add("Text", "w760 cBlue", "Updated files (" allUpdatedFiles.Length "):")
        lvUpdated := updateGui.Add(
            "ListView",
            "w760 r8 Checked -Multi +Background" listColor,
            ["Filename", "Path"])
        for item in allUpdatedFiles {
            SplitPath(item.data.relPath, &displayName)
            lvUpdated.Add(item.shouldCheck ? "Check" : "", displayName, item.data.relPath)
        }
        lvUpdated.ModifyCol(1, 220)
        lvUpdated.ModifyCol(2, 540)
    }

    if newFiles.Length > 0 {
        updateGui.Add("Text", "w760 cGreen y+8", "New files (" newFiles.Length "):")
        lvNew := updateGui.Add(
            "ListView",
            "w760 r8 Checked -Multi +Background" listColor,
            ["Filename", "Path"])
        for item in newFiles {
            SplitPath(item.relPath, &displayName)
            lvNew.Add("Check", displayName, item.relPath)
        }
        lvNew.ModifyCol(1, 220)
        lvNew.ModifyCol(2, 540)
    }

    ; Omitted entirely if neither updater file changed upstream.
    AddUpdaterSnapshotSection(updateGui, updaterState, "Updater snapshots to save")

    if hotstringLibUpdate != "" {
        updateGui.Add("Text", "w760 cBlack y+8", "HotstringLib Update Available")
        radAppend := updateGui.Add(
            "Radio",
            "w760 h38 Checked vHotstringLibMode",
            "Save new library as: Core\" hotstringLibName ". Your current AutoCorrectHotstrings.ahk will not be modified. Use UniqueStringExtractor tool to compare and merge.")
        radReplace := updateGui.Add(
            "Radio",
            "w760 h38",
            "WARNING: Replaces your existing AutoCorrectHotstrings.ahk file. Any customizations you've made will be lost, but you won't have to bother merging the old and new versions.")
        lvHotstringLib := updateGui.Add(
            "ListView",
            "w760 r2 Checked -Multi +Background" listColor " y+5",
            ["Filename", "Path"])
        lvHotstringLib.Add("Check", hotstringLibName, "Core\" hotstringLibName)
        lvHotstringLib.ModifyCol(1, 220)
        lvHotstringLib.ModifyCol(2, 540)
        updateGui.radReplace := radReplace
        radAppend.OnEvent("Click", HotstringLibModeChanged.Bind(updateGui, hotstringLibName, 1))
        radReplace.OnEvent("Click", HotstringLibModeChanged.Bind(updateGui, hotstringLibName, 0))
    }

    updateGui.lvUpdated := lvUpdated
    updateGui.lvNew := lvNew
    updateGui.lvHotstringLib := lvHotstringLib

    ; Only create selection helpers when there is actually selectable work.
    if lvUpdated != "" or lvNew != "" or lvHotstringLib != "" {
        AddSelectionButtons(
            updateGui,
            SetMainUpdateChecks.Bind(updateGui, true),
            SetMainUpdateChecks.Bind(updateGui, false))
    }

    updateGui.Add("Button", "xm w100 y+18", "Update").OnEvent(
        "Click",
        (*) => PerformUpdate(
            updatedFiles,
            rarelyUpdatedFiles,
            newFiles,
            updaterSpecialUpdates,
            updaterState,
            hotstringLibUpdate,
            installDir,
            srcRoot,
            updateGui,
            lvUpdated,
            lvNew,
            lvHotstringLib,
            latestInfo))
    updateGui.Add("Button", "x+10 w120", "Open Debug Log").OnEvent("Click", OpenDebugLog)
    updateGui.Add("Button", "x+10 w100", "Cancel").OnEvent("Click", (*) => ExitApp())
    updateGui.Show()
}
HotstringLibModeChanged(updateGui, hotstringLibName, mode, GuiCtrlObj, Info) {
    ; Change only the displayed destination while preserving the current check state.
    if updateGui.lvHotstringLib != "" {
        wasChecked := updateGui.lvHotstringLib.GetNext(0, "C") = 1
        rowOptions := wasChecked ? "Check" : ""
        updateGui.lvHotstringLib.Delete(1)

        if mode = 1 {
            updateGui.lvHotstringLib.Add(rowOptions, hotstringLibName, "Core\" hotstringLibName)
            LogDebug("Radio mode changed to: MERGE (configured upstream copy)")
        } else {
            updateGui.lvHotstringLib.Add(rowOptions, "AutoCorrectHotstrings.ahk", "Core\AutoCorrectHotstrings.ahk")
            LogDebug("Radio mode changed to: REPLACE")
        }
    }

    if updateGui.radReplace != "" {
        if mode = 0 {
            updateGui.radReplace.Opt("+cFF0000")
            LogDebug("Replace mode selected - radio button text is RED")
        } else {
            updateGui.radReplace.Opt("+" updateGui.fontColor)
            LogDebug("Merge mode selected - radio button text is normal color")
        }
    }
}

PerformUpdate(updatedFiles, rarelyUpdatedFiles, newFiles, updaterSpecialUpdates, updaterState, hotstringLibUpdate, installDir, srcRoot, updateGui, lvUpdated, lvNew, lvHotstringLib, latestInfo) {
    ; First attempt for a fresh cycle. Generic rows are user-selectable; every pending
    ; updater snapshot is a safe read-only operation and is processed automatically.
    LogDebug("Starting initial update operation...")
    formValues := updateGui.Submit(0)
    checkedUpdated := GetCheckedRowSet(lvUpdated)
    checkedNew := GetCheckedRowSet(lvNew)
    hotstringChecked := lvHotstringLib != "" and lvHotstringLib.GetNext(0, "C") = 1

    successfulUpdated := Map()
    successfulNew := Map()
    hotstringSucceeded := false
    hotstringPendingKey := ""
    failures := []
    updatedCycleFiles := []

    for rowNum, _ in checkedUpdated {
        if rowNum <= updatedFiles.Length {
            file := updatedFiles[rowNum]
        } else {
            rarelyIndex := rowNum - updatedFiles.Length
            if rarelyIndex > rarelyUpdatedFiles.Length
                continue
            file := rarelyUpdatedFiles[rarelyIndex]
        }

        result := CopyFileWithDirCreation(file.srcPath, file.path)
        if result.skipped {
            result.displayPath := file.relPath
            failures.Push(result)
        } else {
            successfulUpdated[rowNum] := true
            UpsertUpdatedFileRecord(updatedCycleFiles, file.relPath, A_Now, "Updated")
        }
    }

    for rowNum, _ in checkedNew {
        if rowNum > newFiles.Length
            continue
        file := newFiles[rowNum]
        result := CopyFileWithDirCreation(file.srcPath, file.path)
        if result.skipped {
            result.displayPath := file.relPath
            failures.Push(result)
        } else {
            successfulNew[rowNum] := true
            UpsertUpdatedFileRecord(updatedCycleFiles, file.relPath, A_Now, "Added")
        }
    }

    ; Process every changed updater snapshot automatically, independently.
    for file in updaterSpecialUpdates {
        updaterKind := file.updaterKind
        if updaterState[updaterKind]["status"] != "pending"
            continue
        result := CopyFileWithDirCreation(file.srcPath, file.path)
        if result.skipped {
            result.displayPath := file.destRelPath
            failures.Push(result)
        } else {
            MarkUpdaterCompleted(updaterState, updaterKind, A_Now)
            LogDebug("Saved upstream updater snapshot: " file.destRelPath)
        }
    }

    if lvHotstringLib != "" and hotstringChecked {
        srcFile := hotstringLibUpdate.srcPath

        if formValues.HotstringLibMode = 1 {
            destFile := hotstringLibUpdate.path
            displayPath := hotstringLibUpdate.relPath
            result := CopyFileWithDirCreation(srcFile, destFile)
            if result.skipped {
                result.displayPath := displayPath
                failures.Push(result)
                hotstringPendingKey := MakePendingItem(
                    "Core\AutoCorrectHotstrings.ahk",
                    hotstringLibUpdate.relPath,
                    hotstringLibUpdate.isNew ? "new" : "update")
            } else {
                hotstringSucceeded := true
                UpsertUpdatedFileRecord(updatedCycleFiles, displayPath, A_Now, "Upstream snapshot saved")
            }
        } else {
            activeRelPath := "Core\AutoCorrectHotstrings.ahk"
            activeDest := installDir "\" activeRelPath
            snapshotRelPath := hotstringLibUpdate.relPath
            snapshotDest := hotstringLibUpdate.path

            activeResult := CopyFileWithDirCreation(srcFile, activeDest)
            if activeResult.skipped {
                activeResult.displayPath := activeRelPath
                failures.Push(activeResult)
                hotstringPendingKey := MakeHotstringReplacePendingItem(
                    "Core\AutoCorrectHotstrings.ahk",
                    activeRelPath,
                    snapshotRelPath)
            } else {
                UpsertUpdatedFileRecord(updatedCycleFiles, activeRelPath, A_Now, "Replaced")
                snapshotExistedBefore := FileExist(snapshotDest)
                snapshotResult := CopyFileWithDirCreation(srcFile, snapshotDest)
                if snapshotResult.skipped {
                    snapshotResult.displayPath := snapshotRelPath
                    failures.Push(snapshotResult)
                    hotstringPendingKey := MakePendingItem(
                        "Core\AutoCorrectHotstrings.ahk",
                        snapshotRelPath,
                        snapshotExistedBefore ? "update" : "new")
                } else {
                    hotstringSucceeded := true
                    UpsertUpdatedFileRecord(updatedCycleFiles, snapshotRelPath, A_Now, "Upstream snapshot saved")
                }
            }
        }
    }

    remainingPending := []
    retryCheckedPending := Map()

    if lvUpdated != "" {
        loop lvUpdated.GetCount() {
            if successfulUpdated.Has(A_Index)
                continue
            relPath := lvUpdated.GetText(A_Index, 2)
            if relPath != "" and !ShouldExcludeFromUpdates(relPath) {
                pendingKey := MakePendingItem(relPath, relPath, "update")
                remainingPending.Push(pendingKey)
                if checkedUpdated.Has(A_Index)
                    retryCheckedPending[pendingKey] := true
            }
        }
    }

    if lvNew != "" {
        loop lvNew.GetCount() {
            if successfulNew.Has(A_Index)
                continue
            relPath := lvNew.GetText(A_Index, 2)
            if relPath != "" and !ShouldExcludeFromUpdates(relPath) {
                pendingKey := MakePendingItem(relPath, relPath, "new")
                remainingPending.Push(pendingKey)
                if checkedNew.Has(A_Index)
                    retryCheckedPending[pendingKey] := true
            }
        }
    }

    if lvHotstringLib != "" and !hotstringSucceeded {
        if hotstringPendingKey = "" {
            hotstringPendingKey := (
                formValues.HotstringLibMode = 1
                ? MakePendingItem(
                    "Core\AutoCorrectHotstrings.ahk",
                    hotstringLibUpdate.relPath,
                    hotstringLibUpdate.isNew ? "new" : "update")
                : MakeHotstringReplacePendingItem(
                    "Core\AutoCorrectHotstrings.ahk",
                    "Core\AutoCorrectHotstrings.ahk",
                    hotstringLibUpdate.relPath))
        }
        remainingPending.Push(hotstringPendingKey)
        if hotstringChecked
            retryCheckedPending[hotstringPendingKey] := true
    }

    StoreUpdateCheckInfo(latestInfo["sha"], latestInfo["date"], remainingPending, updatedCycleFiles)
    StoreUpdaterState(updaterState)

    if failures.Length > 0 {
        newState := Map(
            "sha", latestInfo["sha"],
            "date", latestInfo["date"],
            "pendingFiles", remainingPending,
            "updatedFiles", updatedCycleFiles)
        guiFontSize := updateGui.fontSize
        guiFormColor := updateGui.formColor
        guiFontColor := updateGui.fontColor
        guiListColor := updateGui.listColor
        updateGui.Destroy()

        ShowPendingGui(
            remainingPending,
            newState,
            updaterState,
            srcRoot,
            guiFontSize,
            guiFormColor,
            guiFontColor,
            guiListColor,
            retryCheckedPending)
        ShowCopyFailureDialog(failures, guiFontSize, guiFormColor, guiFontColor, guiListColor)
        return
    }

    updateGui.Hide()
    unfinishedCount := remainingPending.Length + GetUpdaterPendingCount(updaterState)
    if unfinishedCount > 0 {
        AcMsgBox.Show(
            "Selected files were updated successfully.`n`n"
            unfinishedCount " pending item(s) remain for later review.",
            "AutoCorrect2 Updater")
    } else {
        AcMsgBox.Show("AutoCorrect2 has been updated successfully!", "AutoCorrect2 Updater")
    }
    ExitApp
}
CopyFileWithDirCreation(srcFile, destPath) {
    SplitPath(destPath, &fileName, &outDir)
    if !DirExist(outDir)
        DirCreate(outDir)

    try {
        FileCopy(srcFile, destPath, 1)
        return {
            skipped: false,
            file: fileName,
            source: srcFile,
            destination: destPath,
            reason: ""
        }
    } catch Error as e {
        ; Capture A_LastError immediately because subsequent OS/file calls can replace it.
        windowsErrorCode := A_LastError
        windowsErrorMessage := ""
        if windowsErrorCode != 0 {
            try windowsErrorMessage := OSError(windowsErrorCode).Message
        } else {
            windowsErrorMessage := "No Windows error code was reported."
        }

        LogDebug(
            "FileCopy failed for '" fileName "': " e.Message
            "; Windows error " windowsErrorCode ": " windowsErrorMessage
            "; FileCopy failure count=" e.Extra)

        return {
            skipped: true,
            file: fileName,
            source: srcFile,
            destination: destPath,
            reason: e.Message,
            windowsErrorCode: windowsErrorCode,
            windowsErrorMessage: windowsErrorMessage
        }
    }
}

IsRarelyUpdated(relPath, folders) {
    for folder in folders
        if SubStr(relPath, 1, StrLen(folder)) = folder
            return true
    return false
}

NormalizePath(path) {
    path := StrReplace(path, "/", "\")
    while SubStr(path, -0) = "\"
        path := SubStr(path, 1, -1)
    return StrLower(path)
}
