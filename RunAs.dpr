(*******************************************************************************

  Jean-Pierre LESUEUR (@DarkCoderSc)
  https://www.phrozen.io/
  jplesueur@phrozen.io

  License : MIT

*******************************************************************************)

program RunAs;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  Windows,
  System.SysUtils,
  UntFunctions in 'Units\UntFunctions.pas';

{-------------------------------------------------------------------------------
  Usage Banner
-------------------------------------------------------------------------------}
function DisplayHelpBanner() : String;
begin
  result := '';
  ///

  WriteLn;
  WriteColoredWord('RunAs');
  Write('.');
  WriteColoredWord('exe');
  WriteLn(' -u <username> -p <password> -e <executable> [-d <domain>] [-a <arguments>] [-h]');
  WriteLn;
  WriteLn('-h : Start process hidden.');
  WriteLn;
end;


var SET_USERNAME   : String  = '';
    SET_PASSWORD   : String  = '';
    SET_DOMAINNAME : String  = '';
    SET_PROGRAM    : String  = '';
    SET_ARGUMENTS  : String  = '';
    SET_HIDDEN     : Boolean = False;

    LRet           : Integer;

begin
  try
    {
      Parse Arguments
    }
    if NOT GetCommandLineOption('u', SET_USERNAME) then
      raise Exception.Create('');

    if NOT GetCommandLineOption('p', SET_PASSWORD) then
      raise Exception.Create('');

    if NOT GetCommandLineOption('e', SET_PROGRAM) then
      raise Exception.Create('');

    GetCommandLineOption('d', SET_DOMAINNAME);
    GetCommandLineOption('a', SET_ARGUMENTS);

    SET_HIDDEN := CommandLineOptionExists('h');

    {
      Run Program
    }
    LRet := CreateProcessAsUser(
                                  SET_PROGRAM,
                                  SET_ARGUMENTS,
                                  SET_USERNAME,
                                  SET_PASSWORD,
                                  SET_DOMAINNAME,
                                  (NOT SET_HIDDEN)
    );

    case LRet of
      {
        Access Denied
      }
      5 : begin
        Debug('Generally this error is related to file access permission. You should place the file you are trying to execute into a folder accessible by any user. (Ex: C:\ProgramData\)', dlWarning);
      end;
    end;
  except
    on E: Exception do begin
      if (E.Message <> '') then
        Debug(Format('%s : %s', [E.ClassName, E.Message]), dlError)
      else
        DisplayHelpBanner();
    end;
  end;
end.
