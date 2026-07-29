-- Source form for the compiled AppleScriptTask file installed as
-- ~/Library/Application Scripts/com.microsoft.Word/VisualTeXWord.scpt

use framework "Foundation"
use framework "AppKit"
use scripting additions

property runtimeSuffix : "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime"
property maximumRelativePathLength : 1024

on OpenVisualTeXSession(sessionId)
    try
        set safeSessionId to my validateSessionId(sessionId as text)
        set visualTeXURL to "visualtex://office/open?session=" & safeSessionId
        set targetURL to current application's NSURL's URLWithString:visualTeXURL
        if targetURL is missing value then error "VisualTeX Session URL is invalid" number 7107
        set openedURL to ((current application's NSWorkspace's sharedWorkspace())'s openURL:targetURL) as boolean
        if not openedURL then error "VisualTeX Session URL could not be opened" number 7108
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end OpenVisualTeXSession

on OpenVisualTeXApplication(ignoredValue)
    try
        set workspaceObject to current application's NSWorkspace's sharedWorkspace()
        set targetURL to workspaceObject's URLForApplicationWithBundleIdentifier:"com.visualtex.studio"
        if targetURL is missing value then error "VisualTeX application is not installed" number 7109
        set openedURL to (workspaceObject's openURL:targetURL) as boolean
        if not openedURL then error "VisualTeX application could not be opened" number 7110
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end OpenVisualTeXApplication

on EnsureVisualTeXDirectory(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        my ensureDirectory(targetPath)
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end EnsureVisualTeXDirectory

on WriteVisualTeXFile(argumentText)
    try
        set {relativePath, encodedData} to my splitPair(argumentText as text)
        set targetPath to my absoluteRuntimePath(relativePath)
        set parentPath to ((current application's NSString's stringWithString:targetPath)'s stringByDeletingLastPathComponent()) as text
        my ensureDirectory(parentPath)
        set normalizedData to my normalizeBase64Url(encodedData)
        set decodedData to current application's NSData's alloc()'s initWithBase64EncodedString:normalizedData options:0
        if decodedData is missing value then error "VisualTeX file bridge Base64URL payload is invalid" number 7125
        set writeSucceeded to (decodedData's writeToFile:targetPath atomically:true) as boolean
        if not writeSucceeded then error "VisualTeX runtime file could not be written" number 7126
        set fileManager to current application's NSFileManager's defaultManager()
        set fileAttributes to current application's NSDictionary's dictionaryWithObject:384 forKey:(current application's NSFilePosixPermissions)
        set attributeError to reference
        set attributesApplied to (fileManager's setAttributes:fileAttributes ofItemAtPath:targetPath |error|:attributeError) as boolean
        if not attributesApplied then error "VisualTeX runtime file permissions could not be applied" number 7127
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end WriteVisualTeXFile

on ReadVisualTeXFile(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        set fileManager to current application's NSFileManager's defaultManager()
        if not ((fileManager's fileExistsAtPath:targetPath) as boolean) then error "VisualTeX runtime file does not exist" number 7128
        set fileData to current application's NSData's dataWithContentsOfFile:targetPath
        if fileData is missing value then error "VisualTeX runtime file could not be read" number 7129
        set encodedData to (fileData's base64EncodedStringWithOptions:0) as text
        set encodedData to my replaceText(encodedData, "+", "-")
        set encodedData to my replaceText(encodedData, "/", "_")
        repeat while encodedData ends with "="
            if (count characters of encodedData) is 1 then
                set encodedData to ""
            else
                set encodedData to text 1 thru -2 of encodedData
            end if
        end repeat
        return "ok|" & encodedData
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end ReadVisualTeXFile

on VisualTeXFileExists(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        set fileManager to current application's NSFileManager's defaultManager()
        if ((fileManager's fileExistsAtPath:targetPath) as boolean) then
            return "ok|1"
        end if
        return "ok|0"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end VisualTeXFileExists

on DeleteVisualTeXFile(relativePath)
    try
        set targetPath to my absoluteRuntimePath(relativePath as text)
        set fileManager to current application's NSFileManager's defaultManager()
        if not ((fileManager's fileExistsAtPath:targetPath) as boolean) then return "ok|1"
        set deleteError to reference
        set deleted to (fileManager's removeItemAtPath:targetPath |error|:deleteError) as boolean
        if not deleted then error "VisualTeX runtime file could not be deleted" number 7130
        return "ok|1"
    on error errorMessage number errorNumber
        return my errorResponse(errorNumber, errorMessage)
    end try
end DeleteVisualTeXFile

on absoluteRuntimePath(relativePath)
    set safeRelativePath to my validateRelativePath(relativePath)
    set rootPath to my ensureRuntimeRoot()
    return rootPath & "/" & safeRelativePath
end absoluteRuntimePath

on ensureRuntimeRoot()
    set homePath to (current application's NSHomeDirectory()) as text
    set rootPath to homePath & "/" & runtimeSuffix
    my ensureDirectory(rootPath)
    return rootPath
end ensureRuntimeRoot

on ensureDirectory(targetPath)
    set fileManager to current application's NSFileManager's defaultManager()
    set directoryError to reference
    set directoryAttributes to current application's NSDictionary's dictionaryWithObject:448 forKey:(current application's NSFilePosixPermissions)
    set created to (fileManager's createDirectoryAtPath:targetPath withIntermediateDirectories:true attributes:directoryAttributes |error|:directoryError) as boolean
    if not created then error "VisualTeX runtime directory could not be created" number 7131
end ensureDirectory

on validateRelativePath(candidate)
    set candidate to candidate as text
    if candidate is "" then error "VisualTeX runtime path is empty" number 7120
    if (count characters of candidate) > maximumRelativePathLength then error "VisualTeX runtime path is too long" number 7121
    if candidate starts with "/" or candidate contains ".." or candidate contains "//" then error "VisualTeX runtime path is unsafe" number 7122
    set allowedCharacters to "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-/"
    repeat with currentCharacter in characters of candidate
        if allowedCharacters does not contain (currentCharacter as text) then error "VisualTeX runtime path contains an unsupported character" number 7123
    end repeat
    return candidate
end validateRelativePath

on splitPair(value)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to "|"
    set fields to text items of value
    set AppleScript's text item delimiters to previousDelimiters
    if (count fields) is not 2 then error "VisualTeX file bridge payload is invalid" number 7124
    return {item 1 of fields, item 2 of fields}
end splitPair

on normalizeBase64Url(encodedData)
    set normalizedData to my replaceText(encodedData as text, "-", "+")
    set normalizedData to my replaceText(normalizedData, "_", "/")
    set remainderValue to (count characters of normalizedData) mod 4
    if remainderValue is 1 then error "VisualTeX file bridge Base64URL payload is invalid" number 7125
    if remainderValue is 2 then set normalizedData to normalizedData & "=="
    if remainderValue is 3 then set normalizedData to normalizedData & "="
    return normalizedData
end normalizeBase64Url

on replaceText(sourceText, searchText, replacementText)
    set previousDelimiters to AppleScript's text item delimiters
    set AppleScript's text item delimiters to searchText
    set sourceItems to text items of sourceText
    set AppleScript's text item delimiters to replacementText
    set resultText to sourceItems as text
    set AppleScript's text item delimiters to previousDelimiters
    return resultText
end replaceText

on validateSessionId(candidate)
    if (count characters of candidate) is not 36 then error "Invalid VisualTeX Session id" number 7101
    if character 9 of candidate is not "-" or character 14 of candidate is not "-" or character 19 of candidate is not "-" or character 24 of candidate is not "-" then error "Invalid VisualTeX Session id" number 7102
    if character 15 of candidate is not "4" then error "Invalid VisualTeX Session version" number 7103
    if "89ab" does not contain character 20 of candidate then error "Invalid VisualTeX Session variant" number 7104

    set allowedHex to "0123456789abcdef"
    repeat with characterIndex from 1 to 36
        set currentCharacter to character characterIndex of candidate
        if characterIndex is 9 or characterIndex is 14 or characterIndex is 19 or characterIndex is 24 then
            if currentCharacter is not "-" then error "Invalid VisualTeX Session id" number 7105
        else if allowedHex does not contain currentCharacter then
            error "Invalid VisualTeX Session id" number 7106
        end if
    end repeat
    return candidate
end validateSessionId

on errorResponse(errorNumber, errorMessage)
    return "error|" & (errorNumber as text) & "|" & my safeError(errorMessage)
end errorResponse

on safeError(value)
    set cleanValue to value as text
    set AppleScript's text item delimiters to {return, linefeed, "|"}
    set cleanItems to text items of cleanValue
    set AppleScript's text item delimiters to " "
    set cleanValue to cleanItems as text
    set AppleScript's text item delimiters to ""
    if (count characters of cleanValue) > 240 then set cleanValue to text 1 thru 240 of cleanValue
    return cleanValue
end safeError
