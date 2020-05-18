(*******************************************************************************

  Jean-Pierre LESUEUR (@DarkCoderSc)
  https://www.phrozen.io/
  jplesueur@phrozen.io

  License : MIT

*******************************************************************************)

unit UntFunctions;

interface

uses Windows, SysUtils;

type
  TDebugLevel = (
                  dlInfo,
                  dlSuccess,
                  dlWarning,
                  dlError
  );

function CreateProcessAsUser(AProgram : String; ACommandLine : String; AUserName, APassword : String; ADomain : String = ''; AVisible : Boolean = True) : Integer;
procedure DumpLastError(APrefix : String = '');
procedure Debug(AMessage : String; ADebugLevel : TDebugLevel = dlInfo);
function GetCommandLineOption(AOption : String; var AValue : String; ACommandLine : String = '') : Boolean; overload;
function GetCommandLineOption(AOption : String; var AValue : String; var AOptionExists : Boolean; ACommandLine : String = '') : Boolean; overload;
function UpdateConsoleAttributes(AConsoleAttributes : Word) : Word;
procedure WriteColoredWord(AString : String);
function CommandLineOptionExists(AOption : String; ACommandLine : String = '') : Boolean;

implementation

{-------------------------------------------------------------------------------
  Check if commandline option is set
-------------------------------------------------------------------------------}
function CommandLineOptionExists(AOption : String; ACommandLine : String = '') : Boolean;
var ADummy : String;
begin
  GetCommandLineOption(AOption, ADummy, result, ACommandLine);
end;

{-------------------------------------------------------------------------------
  Write colored word(s) on current console
-------------------------------------------------------------------------------}
procedure WriteColoredWord(AString : String);
var AOldAttributes : Word;
begin
  AOldAttributes := UpdateConsoleAttributes(FOREGROUND_INTENSITY or FOREGROUND_GREEN);

  Write(AString);

  UpdateConsoleAttributes(AOldAttributes);
end;

{-------------------------------------------------------------------------------
  Update Console Attributes (Changing color for example)

  Returns previous attributes.
-------------------------------------------------------------------------------}
function UpdateConsoleAttributes(AConsoleAttributes : Word) : Word;
var AConsoleHandle        : THandle;
    AConsoleScreenBufInfo : TConsoleScreenBufferInfo;
    b                     : Boolean;
begin
  result := 0;
  ///

  AConsoleHandle := GetStdHandle(STD_OUTPUT_HANDLE);
  if (AConsoleHandle = INVALID_HANDLE_VALUE) then
    Exit();
  ///

  b := GetConsoleScreenBufferInfo(AConsoleHandle, AConsoleScreenBufInfo);

  if b then begin
    SetConsoleTextAttribute(AConsoleHandle, AConsoleAttributes);

    ///
    result := AConsoleScreenBufInfo.wAttributes;
  end;
end;

{-------------------------------------------------------------------------------
  Command Line Parser

  AOption       : Search for specific option Ex: -c.
  AValue        : Next argument string if option is found.
  AOptionExists : Set to true if option is found in command line string.
  ACommandLine  : Command Line String to parse, by default, actual program command line.
-------------------------------------------------------------------------------}
function GetCommandLineOption(AOption : String; var AValue : String; var AOptionExists : Boolean; ACommandLine : String = '') : Boolean;
var ACount    : Integer;
    pElements : Pointer;
    I         : Integer;
    ACurArg   : String;
    hShell32  : THandle;

    CommandLineToArgvW : function(lpCmdLine : LPCWSTR; var pNumArgs : Integer) : LPWSTR; stdcall;
type
  TArgv = array[0..0] of PWideChar;
begin
  result := False;
  ///

  AOptionExists := False;

  hShell32 := LoadLibrary('SHELL32.DLL');
  if (hShell32 = 0) then begin
    Debug('Could load Shell32.dll Library.', dlError);
    Exit();
  end;
  try
    @CommandLineToArgvW := GetProcAddress(hShell32, 'CommandLineToArgvW');
    if NOT Assigned(CommandLineToArgvW) then begin
      Debug('Could load CommandLineToArgvW API.', dlError);
      Exit();
    end;

    if (ACommandLine = '') then begin
      ACommandLine := GetCommandLineW();
    end;

    pElements := CommandLineToArgvW(PWideChar(ACommandLine), ACount);

    if NOT Assigned(pElements) then
      Exit();

    AOption := '-' + AOption;

    if (Length(AOption) > 2) then
      AOption := '-' + AOption;

    for I := 0 to ACount -1 do begin
      ACurArg := UnicodeString((TArgv(pElements^)[I]));
      ///

      if (ACurArg <> AOption) then
        continue;

      AOptionExists := True;

      // Retrieve Next Arg
      if I <> (ACount -1) then begin
        AValue := UnicodeString((TArgv(pElements^)[I+1]));

        ///
        result := True;
      end;
    end;
  finally
    FreeLibrary(hShell32);
  end;
end;

function GetCommandLineOption(AOption : String; var AValue : String; ACommandLine : String = '') : Boolean;
var AExists : Boolean;
begin
  result := GetCommandLineOption(AOption, AValue, AExists, ACommandLine);
end;

{-------------------------------------------------------------------------------
  Spawn a process as another user.
-------------------------------------------------------------------------------}
function CreateProcessAsUser(AProgram : String; ACommandLine : String; AUserName, APassword : String; ADomain : String = ''; AVisible : Boolean = True) : Integer;
var AStartupInfo : TStartupInfo;
    AProcessInfo : TProcessInformation;
    hAdvApi32    : THandle;
    b            : Boolean;

    CreateProcessWithLogonW : function(
                                          lpUsername, lpDomain, lpPassword: LPCWSTR;
                                          dwLogonFlags: DWORD;
                                          lpApplicationName: LPCWSTR;
                                          lpCommandLine: LPWSTR;
                                          dwCreationFlags: DWORD;
                                          lpEnvironment: LPVOID;
                                          lpCurrentDirectory: LPCWSTR;
                                          const lpStartupInfo: STARTUPINFOW;
                                          var lpProcessInformation: PROCESS_INFORMATION
                                      ): BOOL; stdcall;
begin
  result := 0;
  ///

  hAdvapi32 := LoadLibrary('ADVAPI32.DLL');
  if (hAdvapi32 = 0) then begin
    Debug('Could load Advapi32.dll Library.', dlError);

    Exit();
  end;
  try
    @CreateProcessWithLogonW := GetProcAddress(hAdvapi32, 'CreateProcessWithLogonW');

    if NOT Assigned(CreateProcessWithLogonW) then begin
      Debug('Could load CreateProcessWithLogonW API.', dlError);
      Exit();
    end;
    ///

    if (ADomain = '') then
      ADomain := GetEnvironmentVariable('USERDOMAIN');
    ///

    ACommandLine := Format('%s %s', [AProgram, ACommandLine]);

    UniqueString(ACommandLine);
    UniqueString(AUserName);
    UniqueString(APassword);
    UniqueString(ADomain);
    ///

    ZeroMemory(@AProcessInfo, SizeOf(TProcessInformation));
    ZeroMemory(@AStartupInfo, Sizeof(TStartupInfo));

    AStartupInfo.cb          := SizeOf(TStartupInfo);

    if AVisible then
      AStartupInfo.wShowWindow := SW_SHOW
    else
      AStartupInfo.wShowWindow := SW_HIDE;

    AStartupInfo.dwFlags     := (STARTF_USESHOWWINDOW);

    b := CreateProcessWithLogonW(
                                       PWideChar(AUserName),
                                       PWideChar(ADomain),
                                       PWideChar(APassword),
                                       0,
                                       nil,
                                       PWideChar(ACommandLine),
                                       0,
                                       nil,
                                       nil,
                                       AStartupInfo,
                                       AProcessInfo
    );

    result := GetLastError();

    if (NOT b) then
      DumpLastError('CreateProcessWithLogonW')
    else
      Debug(Format('Process spawned as user=[%s], ProcessId=[%d] and ProcessHandle=[%d].', [AUserName, AProcessInfo.dwProcessId, AProcessInfo.hProcess]), dlSuccess);
  finally
    FreeLibrary(hAdvApi32);
  end;
end;

{-------------------------------------------------------------------------------
  Debug Defs
-------------------------------------------------------------------------------}
procedure Debug(AMessage : String; ADebugLevel : TDebugLevel = dlInfo);
var AConsoleHandle        : THandle;
    AConsoleScreenBufInfo : TConsoleScreenBufferInfo;
    b                     : Boolean;
    AStatus               : String;
    AColor                : Integer;
begin
  AConsoleHandle := GetStdHandle(STD_OUTPUT_HANDLE);
  if (AConsoleHandle = INVALID_HANDLE_VALUE) then
    Exit();
  ///

  b := GetConsoleScreenBufferInfo(AConsoleHandle, AConsoleScreenBufInfo);

  case ADebugLevel of
    dlSuccess : begin
      AStatus := #32 + 'OK' + #32;
      AColor  := FOREGROUND_GREEN;
    end;

    dlWarning : begin
      AStatus := #32 + '!!' + #32;
      AColor  := (FOREGROUND_RED or FOREGROUND_GREEN);
    end;

    dlError : begin
      AStatus := #32 + 'KO' + #32;
      AColor  := FOREGROUND_RED;
    end;

    else begin
      AStatus := 'INFO';
      AColor  := FOREGROUND_BLUE;
    end;
  end;

  Write('[');
  if b then
    b := SetConsoleTextAttribute(AConsoleHandle, FOREGROUND_INTENSITY or (AColor));
  try
    Write(AStatus);
  finally
    if b then
      SetConsoleTextAttribute(AConsoleHandle, AConsoleScreenBufInfo.wAttributes);
  end;
  Write(']' + #32);

  ///
  WriteLn(AMessage);
end;

procedure DumpLastError(APrefix : String = '');
var ACode         : Integer;
    AFinalMessage : String;
begin
  ACode := GetLastError();

  AFinalMessage := '';

  if (ACode <> 0) then begin
    AFinalMessage := Format('Error_Msg=[%s], Error_Code=[%d]', [SysErrorMessage(ACode), ACode]);

    if (APrefix <> '') then
      AFinalMessage := Format('%s: %s', [APrefix, AFinalMessage]);

    ///
    Debug(AFinalMessage, dlError);
  end;
end;

end.
