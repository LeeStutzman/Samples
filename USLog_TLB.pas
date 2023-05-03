unit USLog_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// $Rev: 8291 $
// File generated on 7/6/2018 10:35:09 AM from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Projects\UltiPro\DLL\USLog\USLog.tlb (1)
// LIBID: {7CB397D9-D997-48A8-8CBD-E6F07AC41F7A}
// LCID: 0
// Helpfile: 
// HelpString: USLog Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
//   (2) v2.6 ADODB, (C:\Program Files (x86)\Common Files\System\ado\msado26.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, ADODB_TLB, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  USLogMajorVersion = 1;
  USLogMinorVersion = 0;

  LIBID_USLog: TGUID = '{7CB397D9-D997-48A8-8CBD-E6F07AC41F7A}';

  IID_IUSLog: TGUID = '{C9F1AFDF-B2B2-4D9D-BB70-5E37E91169F1}';
  IID_IFileLog: TGUID = '{07760051-92AD-4B14-982D-748E79869B25}';
  CLASS_FileLog: TGUID = '{4127204A-BA80-4E79-BDE1-AE2BB33B96F1}';
  DIID_IUSLogEvents: TGUID = '{AA74109E-4F46-404C-A65B-2C4B3E6397A7}';
  IID_ITableLog: TGUID = '{B7FE6EDD-E951-47AB-B37E-01A1A5FBF864}';
  CLASS_TableLog: TGUID = '{5B6ECF6F-D21F-47F2-B2AA-17627E4A3A10}';
  IID_IDelimLog: TGUID = '{EEADB0C0-0FA5-491D-84B0-1AB6369CE16E}';
  CLASS_DelimLog: TGUID = '{D2F8AA7F-C078-4BC8-A5F8-30F6FF5F6C39}';
  IID_IXMLLog: TGUID = '{4ACB08BE-E74C-4466-BA1C-BD5E0453B07B}';
  CLASS_XMLLog: TGUID = '{9ABE9753-E0C9-4C55-901D-5E3451B115D0}';
  IID_ILogFactory: TGUID = '{6EA0A22D-8A92-4E96-9DEB-B82EA5D9B1E3}';
  CLASS_LogFactory: TGUID = '{4EF34798-02DB-45A8-8D1C-A5CA1F0A22F3}';
  CLASS_USLogBase: TGUID = '{5CFB1BAD-9832-429B-A6E1-7FB075914CC9}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum LogLevel
type
  LogLevel = TOleEnum;
const
  llBeginEnd = $00000000;
  llDebug = $00000001;
  llStatus = $00000002;
  llWarning = $00000003;
  llPriority = $00000004;

// Constants for enum LogType
type
  LogType = TOleEnum;
const
  ltFile = $00000000;
  ltDelimited = $00000001;
  ltXML = $00000002;
  ltTable = $00000003;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IUSLog = interface;
  IUSLogDisp = dispinterface;
  IFileLog = interface;
  IFileLogDisp = dispinterface;
  IUSLogEvents = dispinterface;
  ITableLog = interface;
  ITableLogDisp = dispinterface;
  IDelimLog = interface;
  IDelimLogDisp = dispinterface;
  IXMLLog = interface;
  IXMLLogDisp = dispinterface;
  ILogFactory = interface;
  ILogFactoryDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  FileLog = IFileLog;
  TableLog = ITableLog;
  DelimLog = IDelimLog;
  XMLLog = IXMLLog;
  LogFactory = ILogFactory;
  USLogBase = IUSLog;


// *********************************************************************//
// Interface: IUSLog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C9F1AFDF-B2B2-4D9D-BB70-5E37E91169F1}
// *********************************************************************//
  IUSLog = interface(IDispatch)
    ['{C9F1AFDF-B2B2-4D9D-BB70-5E37E91169F1}']
    function Get_LogEnabled: WordBool; safecall;
    procedure Set_LogEnabled(Value: WordBool); safecall;
    function Get_LogName: WideString; safecall;
    procedure Set_LogName(const Value: WideString); safecall;
    function Get_LogDate: WordBool; safecall;
    procedure Set_LogDate(Value: WordBool); safecall;
    function Get_LogDateFormat: WideString; safecall;
    procedure Set_LogDateFormat(const Value: WideString); safecall;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); safecall;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); safecall;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); safecall;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); safecall;
    function StartTimer: Integer; safecall;
    function CheckTimer(Key: Integer): WideString; safecall;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); safecall;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); safecall;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); safecall;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); safecall;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); safecall;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); safecall;
    procedure ClearLog; safecall;
    procedure Lock; safecall;
    procedure Unlock; safecall;
    function Get_LogReference: WordBool; safecall;
    procedure Set_LogReference(Value: WordBool); safecall;
    function Get_LogProcedureInfo: WordBool; safecall;
    procedure Set_LogProcedureInfo(Value: WordBool); safecall;
    function Get_LogPath: WideString; safecall;
    procedure Set_LogPath(const Value: WideString); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const Value: WideString); safecall;
    function Get_LogLevel: LogLevel; safecall;
    procedure Set_LogLevel(Value: LogLevel); safecall;
    function Get_UniqueFileName: WordBool; safecall;
    procedure Set_UniqueFileName(Value: WordBool); safecall;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); safecall;
    property LogEnabled: WordBool read Get_LogEnabled write Set_LogEnabled;
    property LogName: WideString read Get_LogName write Set_LogName;
    property LogDate: WordBool read Get_LogDate write Set_LogDate;
    property LogDateFormat: WideString read Get_LogDateFormat write Set_LogDateFormat;
    property LogReference: WordBool read Get_LogReference write Set_LogReference;
    property LogProcedureInfo: WordBool read Get_LogProcedureInfo write Set_LogProcedureInfo;
    property LogPath: WideString read Get_LogPath write Set_LogPath;
    property ID: WideString read Get_ID write Set_ID;
    property LogLevel: LogLevel read Get_LogLevel write Set_LogLevel;
    property UniqueFileName: WordBool read Get_UniqueFileName write Set_UniqueFileName;
  end;

// *********************************************************************//
// DispIntf:  IUSLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C9F1AFDF-B2B2-4D9D-BB70-5E37E91169F1}
// *********************************************************************//
  IUSLogDisp = dispinterface
    ['{C9F1AFDF-B2B2-4D9D-BB70-5E37E91169F1}']
    property LogEnabled: WordBool dispid 1;
    property LogName: WideString dispid 2;
    property LogDate: WordBool dispid 3;
    property LogDateFormat: WideString dispid 4;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); dispid 5;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); dispid 6;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); dispid 7;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); dispid 8;
    function StartTimer: Integer; dispid 9;
    function CheckTimer(Key: Integer): WideString; dispid 10;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); dispid 11;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 12;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); dispid 13;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); dispid 14;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); dispid 15;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); dispid 16;
    procedure ClearLog; dispid 17;
    procedure Lock; dispid 18;
    procedure Unlock; dispid 19;
    property LogReference: WordBool dispid 20;
    property LogProcedureInfo: WordBool dispid 21;
    property LogPath: WideString dispid 22;
    property ID: WideString dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UniqueFileName: WordBool dispid 25;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 201;
  end;

// *********************************************************************//
// Interface: IFileLog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {07760051-92AD-4B14-982D-748E79869B25}
// *********************************************************************//
  IFileLog = interface(IUSLog)
    ['{07760051-92AD-4B14-982D-748E79869B25}']
    function Get_LogFileName: WideString; safecall;
    procedure Set_LogFileName(const Value: WideString); safecall;
    function Get_UseRunHeader: WordBool; safecall;
    procedure Set_UseRunHeader(Value: WordBool); safecall;
    property LogFileName: WideString read Get_LogFileName write Set_LogFileName;
    property UseRunHeader: WordBool read Get_UseRunHeader write Set_UseRunHeader;
  end;

// *********************************************************************//
// DispIntf:  IFileLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {07760051-92AD-4B14-982D-748E79869B25}
// *********************************************************************//
  IFileLogDisp = dispinterface
    ['{07760051-92AD-4B14-982D-748E79869B25}']
    property LogFileName: WideString dispid 99;
    property UseRunHeader: WordBool dispid 98;
    property LogEnabled: WordBool dispid 1;
    property LogName: WideString dispid 2;
    property LogDate: WordBool dispid 3;
    property LogDateFormat: WideString dispid 4;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); dispid 5;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); dispid 6;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); dispid 7;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); dispid 8;
    function StartTimer: Integer; dispid 9;
    function CheckTimer(Key: Integer): WideString; dispid 10;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); dispid 11;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 12;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); dispid 13;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); dispid 14;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); dispid 15;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); dispid 16;
    procedure ClearLog; dispid 17;
    procedure Lock; dispid 18;
    procedure Unlock; dispid 19;
    property LogReference: WordBool dispid 20;
    property LogProcedureInfo: WordBool dispid 21;
    property LogPath: WideString dispid 22;
    property ID: WideString dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UniqueFileName: WordBool dispid 25;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 201;
  end;

// *********************************************************************//
// DispIntf:  IUSLogEvents
// Flags:     (4096) Dispatchable
// GUID:      {AA74109E-4F46-404C-A65B-2C4B3E6397A7}
// *********************************************************************//
  IUSLogEvents = dispinterface
    ['{AA74109E-4F46-404C-A65B-2C4B3E6397A7}']
    procedure OnLog(const Message: WideString); dispid 1;
  end;

// *********************************************************************//
// Interface: ITableLog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B7FE6EDD-E951-47AB-B37E-01A1A5FBF864}
// *********************************************************************//
  ITableLog = interface(IUSLog)
    ['{B7FE6EDD-E951-47AB-B37E-01A1A5FBF864}']
    function Get_TableName: WideString; safecall;
    procedure Set_TableName(const Value: WideString); safecall;
    function Get_MessageFieldName: WideString; safecall;
    procedure Set_MessageFieldName(const Value: WideString); safecall;
    function Get_ObjectFieldName: WideString; safecall;
    procedure Set_ObjectFieldName(const Value: WideString); safecall;
    function Get_ProcFieldName: WideString; safecall;
    procedure Set_ProcFieldName(const Value: WideString); safecall;
    function Get_ReferenceFieldName: WideString; safecall;
    procedure Set_ReferenceFieldName(const Value: WideString); safecall;
    procedure SetTableInfo(const sTable: WideString; const sMessageField: WideString; 
                           const sObjectField: WideString; const sProcField: WideString; 
                           const sReferenceField: WideString; const sLevelField: WideString; 
                           const sIDField: WideString); safecall;
    function Get_ConnectionString: WideString; safecall;
    procedure Set_ConnectionString(const Value: WideString); safecall;
    function Get_LevelFieldName: WideString; safecall;
    procedure Set_LevelFieldName(const Value: WideString); safecall;
    function Get_IDFieldName: WideString; safecall;
    procedure Set_IDFieldName(const Value: WideString); safecall;
    function Get_ADOConnection: _Connection; safecall;
    procedure Set_ADOConnection(const Value: _Connection); safecall;
    property TableName: WideString read Get_TableName write Set_TableName;
    property MessageFieldName: WideString read Get_MessageFieldName write Set_MessageFieldName;
    property ObjectFieldName: WideString read Get_ObjectFieldName write Set_ObjectFieldName;
    property ProcFieldName: WideString read Get_ProcFieldName write Set_ProcFieldName;
    property ReferenceFieldName: WideString read Get_ReferenceFieldName write Set_ReferenceFieldName;
    property ConnectionString: WideString read Get_ConnectionString write Set_ConnectionString;
    property LevelFieldName: WideString read Get_LevelFieldName write Set_LevelFieldName;
    property IDFieldName: WideString read Get_IDFieldName write Set_IDFieldName;
    property ADOConnection: _Connection read Get_ADOConnection write Set_ADOConnection;
  end;

// *********************************************************************//
// DispIntf:  ITableLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {B7FE6EDD-E951-47AB-B37E-01A1A5FBF864}
// *********************************************************************//
  ITableLogDisp = dispinterface
    ['{B7FE6EDD-E951-47AB-B37E-01A1A5FBF864}']
    property TableName: WideString dispid 99;
    property MessageFieldName: WideString dispid 98;
    property ObjectFieldName: WideString dispid 97;
    property ProcFieldName: WideString dispid 96;
    property ReferenceFieldName: WideString dispid 94;
    procedure SetTableInfo(const sTable: WideString; const sMessageField: WideString; 
                           const sObjectField: WideString; const sProcField: WideString; 
                           const sReferenceField: WideString; const sLevelField: WideString; 
                           const sIDField: WideString); dispid 93;
    property ConnectionString: WideString dispid 92;
    property LevelFieldName: WideString dispid 91;
    property IDFieldName: WideString dispid 90;
    property ADOConnection: _Connection dispid 26;
    property LogEnabled: WordBool dispid 1;
    property LogName: WideString dispid 2;
    property LogDate: WordBool dispid 3;
    property LogDateFormat: WideString dispid 4;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); dispid 5;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); dispid 6;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); dispid 7;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); dispid 8;
    function StartTimer: Integer; dispid 9;
    function CheckTimer(Key: Integer): WideString; dispid 10;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); dispid 11;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 12;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); dispid 13;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); dispid 14;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); dispid 15;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); dispid 16;
    procedure ClearLog; dispid 17;
    procedure Lock; dispid 18;
    procedure Unlock; dispid 19;
    property LogReference: WordBool dispid 20;
    property LogProcedureInfo: WordBool dispid 21;
    property LogPath: WideString dispid 22;
    property ID: WideString dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UniqueFileName: WordBool dispid 25;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 201;
  end;

// *********************************************************************//
// Interface: IDelimLog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EEADB0C0-0FA5-491D-84B0-1AB6369CE16E}
// *********************************************************************//
  IDelimLog = interface(IUSLog)
    ['{EEADB0C0-0FA5-491D-84B0-1AB6369CE16E}']
    function Get_LogFileName: WideString; safecall;
    procedure Set_LogFileName(const Value: WideString); safecall;
    function Get_FieldSeparator: WideString; safecall;
    procedure Set_FieldSeparator(const Value: WideString); safecall;
    function Get_FieldQualifier: WideString; safecall;
    procedure Set_FieldQualifier(const Value: WideString); safecall;
    property LogFileName: WideString read Get_LogFileName write Set_LogFileName;
    property FieldSeparator: WideString read Get_FieldSeparator write Set_FieldSeparator;
    property FieldQualifier: WideString read Get_FieldQualifier write Set_FieldQualifier;
  end;

// *********************************************************************//
// DispIntf:  IDelimLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {EEADB0C0-0FA5-491D-84B0-1AB6369CE16E}
// *********************************************************************//
  IDelimLogDisp = dispinterface
    ['{EEADB0C0-0FA5-491D-84B0-1AB6369CE16E}']
    property LogFileName: WideString dispid 99;
    property FieldSeparator: WideString dispid 98;
    property FieldQualifier: WideString dispid 97;
    property LogEnabled: WordBool dispid 1;
    property LogName: WideString dispid 2;
    property LogDate: WordBool dispid 3;
    property LogDateFormat: WideString dispid 4;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); dispid 5;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); dispid 6;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); dispid 7;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); dispid 8;
    function StartTimer: Integer; dispid 9;
    function CheckTimer(Key: Integer): WideString; dispid 10;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); dispid 11;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 12;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); dispid 13;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); dispid 14;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); dispid 15;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); dispid 16;
    procedure ClearLog; dispid 17;
    procedure Lock; dispid 18;
    procedure Unlock; dispid 19;
    property LogReference: WordBool dispid 20;
    property LogProcedureInfo: WordBool dispid 21;
    property LogPath: WideString dispid 22;
    property ID: WideString dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UniqueFileName: WordBool dispid 25;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 201;
  end;

// *********************************************************************//
// Interface: IXMLLog
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {4ACB08BE-E74C-4466-BA1C-BD5E0453B07B}
// *********************************************************************//
  IXMLLog = interface(IUSLog)
    ['{4ACB08BE-E74C-4466-BA1C-BD5E0453B07B}']
    function Get_LogFileName: WideString; safecall;
    procedure Set_LogFileName(const Value: WideString); safecall;
    property LogFileName: WideString read Get_LogFileName write Set_LogFileName;
  end;

// *********************************************************************//
// DispIntf:  IXMLLogDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {4ACB08BE-E74C-4466-BA1C-BD5E0453B07B}
// *********************************************************************//
  IXMLLogDisp = dispinterface
    ['{4ACB08BE-E74C-4466-BA1C-BD5E0453B07B}']
    property LogFileName: WideString dispid 99;
    property LogEnabled: WordBool dispid 1;
    property LogName: WideString dispid 2;
    property LogDate: WordBool dispid 3;
    property LogDateFormat: WideString dispid 4;
    procedure Log(const Msg: WideString; Level: Integer; const ProcName: WideString; 
                  const ObjectName: WideString; const Reference: WideString); dispid 5;
    procedure LogException(const Msg: WideString; const ProcName: WideString; 
                           const ObjectName: WideString; const Reference: WideString); dispid 6;
    procedure LogProcEnd(const ProcName: WideString; const ObjectName: WideString; 
                         const Reference: WideString); dispid 7;
    procedure LogFuncEnd(const ProcName: WideString; Value: OleVariant; 
                         const ObjectName: WideString; const Reference: WideString); dispid 8;
    function StartTimer: Integer; dispid 9;
    function CheckTimer(Key: Integer): WideString; dispid 10;
    procedure LogTimer(Key: Integer; const Msg: WideString; const Reference: WideString; 
                       Level: Integer); dispid 11;
    procedure LogProcStart(const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 12;
    procedure LogDebug(const Msg: WideString; const ProcName: WideString; 
                       const ObjectName: WideString; const Reference: WideString); dispid 13;
    procedure LogStatus(const Msg: WideString; const ProcName: WideString; 
                        const ObjectName: WideString; const Reference: WideString); dispid 14;
    procedure LogWarning(const Msg: WideString; const ProcName: WideString; 
                         const ObjectName: WideString; const Reference: WideString); dispid 15;
    procedure LogNotImplemented(const ProcName: WideString; const ObjectName: WideString; 
                                const Reference: WideString); dispid 16;
    procedure ClearLog; dispid 17;
    procedure Lock; dispid 18;
    procedure Unlock; dispid 19;
    property LogReference: WordBool dispid 20;
    property LogProcedureInfo: WordBool dispid 21;
    property LogPath: WideString dispid 22;
    property ID: WideString dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UniqueFileName: WordBool dispid 25;
    procedure LogParameter(const ParamName: WideString; ParamValue: OleVariant; 
                           const ProcName: WideString; const ObjectName: WideString; 
                           const Reference: WideString); dispid 201;
  end;

// *********************************************************************//
// Interface: ILogFactory
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EA0A22D-8A92-4E96-9DEB-B82EA5D9B1E3}
// *********************************************************************//
  ILogFactory = interface(IDispatch)
    ['{6EA0A22D-8A92-4E96-9DEB-B82EA5D9B1E3}']
    procedure CopyFrom(const Source: ILogFactory); safecall;
    procedure Clear; safecall;
    function ReadFromReg(iRootKey: LongWord; const sKey: WideString): WordBool; safecall;
    function ReadFromINI(const sFileName: WideString; const sSection: WideString): WordBool; safecall;
    function WriteToReg(iRootKey: LongWord; const sKey: WideString): WordBool; safecall;
    function WriteToINI(const sFileName: WideString; const sSection: WideString): WordBool; safecall;
    function ReadFromXML(const sFileName: WideString): WordBool; safecall;
    function ReadFromXMLStr(const sXML: WideString): WordBool; safecall;
    function CreateLog: IUSLog; safecall;
    function SetupObjFromInfo(const Obj: IUSLog): WordBool; safecall;
    function RegCreateLog(iRootKey: LongWord; const sKey: WideString): IUSLog; safecall;
    function INICreateLog(const sFileName: WideString; const sSection: WideString): IUSLog; safecall;
    function XMLCreateLog(const sFileName: WideString): IUSLog; safecall;
    function XMLStrCreateLog(const sXML: WideString): IUSLog; safecall;
    function Get_LogType: LogType; safecall;
    procedure Set_LogType(Value: LogType); safecall;
    function Get_LogName: WideString; safecall;
    procedure Set_LogName(const Value: WideString); safecall;
    function Get_LogEnabled: WordBool; safecall;
    procedure Set_LogEnabled(Value: WordBool); safecall;
    function Get_LogDate: WordBool; safecall;
    procedure Set_LogDate(Value: WordBool); safecall;
    function Get_LogDateFormat: WideString; safecall;
    procedure Set_LogDateFormat(const Value: WideString); safecall;
    function Get_LogReference: WordBool; safecall;
    procedure Set_LogReference(Value: WordBool); safecall;
    function Get_LogProcInfo: WordBool; safecall;
    procedure Set_LogProcInfo(Value: WordBool); safecall;
    function Get_LogLevel: LogLevel; safecall;
    procedure Set_LogLevel(Value: LogLevel); safecall;
    function Get_UseRunHeader: WordBool; safecall;
    procedure Set_UseRunHeader(Value: WordBool); safecall;
    function Get_FieldSeparator: WideString; safecall;
    procedure Set_FieldSeparator(const Value: WideString); safecall;
    function Get_FieldQualifier: WideString; safecall;
    procedure Set_FieldQualifier(const Value: WideString); safecall;
    function Get_MsgField: WideString; safecall;
    procedure Set_MsgField(const Value: WideString); safecall;
    function Get_ObjField: WideString; safecall;
    procedure Set_ObjField(const Value: WideString); safecall;
    function Get_ProcField: WideString; safecall;
    procedure Set_ProcField(const Value: WideString); safecall;
    function Get_RefField: WideString; safecall;
    procedure Set_RefField(const Value: WideString); safecall;
    function Get_LevelField: WideString; safecall;
    procedure Set_LevelField(const Value: WideString); safecall;
    function Get_LogPath: WideString; safecall;
    procedure Set_LogPath(const Value: WideString); safecall;
    function Get_IDField: WideString; safecall;
    procedure Set_IDField(const Value: WideString); safecall;
    function Get_ID: WideString; safecall;
    procedure Set_ID(const Value: WideString); safecall;
    procedure SetIDByType(const sFileBase: WideString; const sTblBase: WideString); safecall;
    function Get_ConnectionString: WideString; safecall;
    procedure Set_ConnectionString(const Value: WideString); safecall;
    function Get_UniqueFileName: WordBool; safecall;
    procedure Set_UniqueFileName(Value: WordBool); safecall;
    function Get_LogXMLSchema: WideString; safecall;
    function Get_LogSetupXMLSchema: WideString; safecall;
    function WriteToXML(const sFileName: WideString): WordBool; safecall;
    function WriteToXMLStr: WideString; safecall;
    function Get_ErrorMsg: WideString; safecall;
    function Get_LogAppendJobIDToFilename: WordBool; safecall;
    procedure Set_LogAppendJobIDToFilename(Value: WordBool); safecall;
    property LogType: LogType read Get_LogType write Set_LogType;
    property LogName: WideString read Get_LogName write Set_LogName;
    property LogEnabled: WordBool read Get_LogEnabled write Set_LogEnabled;
    property LogDate: WordBool read Get_LogDate write Set_LogDate;
    property LogDateFormat: WideString read Get_LogDateFormat write Set_LogDateFormat;
    property LogReference: WordBool read Get_LogReference write Set_LogReference;
    property LogProcInfo: WordBool read Get_LogProcInfo write Set_LogProcInfo;
    property LogLevel: LogLevel read Get_LogLevel write Set_LogLevel;
    property UseRunHeader: WordBool read Get_UseRunHeader write Set_UseRunHeader;
    property FieldSeparator: WideString read Get_FieldSeparator write Set_FieldSeparator;
    property FieldQualifier: WideString read Get_FieldQualifier write Set_FieldQualifier;
    property MsgField: WideString read Get_MsgField write Set_MsgField;
    property ObjField: WideString read Get_ObjField write Set_ObjField;
    property ProcField: WideString read Get_ProcField write Set_ProcField;
    property RefField: WideString read Get_RefField write Set_RefField;
    property LevelField: WideString read Get_LevelField write Set_LevelField;
    property LogPath: WideString read Get_LogPath write Set_LogPath;
    property IDField: WideString read Get_IDField write Set_IDField;
    property ID: WideString read Get_ID write Set_ID;
    property ConnectionString: WideString read Get_ConnectionString write Set_ConnectionString;
    property UniqueFileName: WordBool read Get_UniqueFileName write Set_UniqueFileName;
    property LogXMLSchema: WideString read Get_LogXMLSchema;
    property LogSetupXMLSchema: WideString read Get_LogSetupXMLSchema;
    property ErrorMsg: WideString read Get_ErrorMsg;
    property LogAppendJobIDToFilename: WordBool read Get_LogAppendJobIDToFilename write Set_LogAppendJobIDToFilename;
  end;

// *********************************************************************//
// DispIntf:  ILogFactoryDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {6EA0A22D-8A92-4E96-9DEB-B82EA5D9B1E3}
// *********************************************************************//
  ILogFactoryDisp = dispinterface
    ['{6EA0A22D-8A92-4E96-9DEB-B82EA5D9B1E3}']
    procedure CopyFrom(const Source: ILogFactory); dispid 1;
    procedure Clear; dispid 2;
    function ReadFromReg(iRootKey: LongWord; const sKey: WideString): WordBool; dispid 3;
    function ReadFromINI(const sFileName: WideString; const sSection: WideString): WordBool; dispid 4;
    function WriteToReg(iRootKey: LongWord; const sKey: WideString): WordBool; dispid 5;
    function WriteToINI(const sFileName: WideString; const sSection: WideString): WordBool; dispid 6;
    function ReadFromXML(const sFileName: WideString): WordBool; dispid 7;
    function ReadFromXMLStr(const sXML: WideString): WordBool; dispid 8;
    function CreateLog: IUSLog; dispid 10;
    function SetupObjFromInfo(const Obj: IUSLog): WordBool; dispid 11;
    function RegCreateLog(iRootKey: LongWord; const sKey: WideString): IUSLog; dispid 12;
    function INICreateLog(const sFileName: WideString; const sSection: WideString): IUSLog; dispid 13;
    function XMLCreateLog(const sFileName: WideString): IUSLog; dispid 14;
    function XMLStrCreateLog(const sXML: WideString): IUSLog; dispid 15;
    property LogType: LogType dispid 16;
    property LogName: WideString dispid 17;
    property LogEnabled: WordBool dispid 18;
    property LogDate: WordBool dispid 19;
    property LogDateFormat: WideString dispid 20;
    property LogReference: WordBool dispid 22;
    property LogProcInfo: WordBool dispid 23;
    property LogLevel: LogLevel dispid 24;
    property UseRunHeader: WordBool dispid 25;
    property FieldSeparator: WideString dispid 26;
    property FieldQualifier: WideString dispid 27;
    property MsgField: WideString dispid 28;
    property ObjField: WideString dispid 29;
    property ProcField: WideString dispid 30;
    property RefField: WideString dispid 32;
    property LevelField: WideString dispid 33;
    property LogPath: WideString dispid 34;
    property IDField: WideString dispid 37;
    property ID: WideString dispid 38;
    procedure SetIDByType(const sFileBase: WideString; const sTblBase: WideString); dispid 40;
    property ConnectionString: WideString dispid 36;
    property UniqueFileName: WordBool dispid 21;
    property LogXMLSchema: WideString readonly dispid 35;
    property LogSetupXMLSchema: WideString readonly dispid 39;
    function WriteToXML(const sFileName: WideString): WordBool; dispid 41;
    function WriteToXMLStr: WideString; dispid 42;
    property ErrorMsg: WideString readonly dispid 9;
    property LogAppendJobIDToFilename: WordBool dispid 201;
  end;

// *********************************************************************//
// The Class CoFileLog provides a Create and CreateRemote method to          
// create instances of the default interface IFileLog exposed by              
// the CoClass FileLog. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFileLog = class
    class function Create: IFileLog;
    class function CreateRemote(const MachineName: string): IFileLog;
  end;

// *********************************************************************//
// The Class CoTableLog provides a Create and CreateRemote method to          
// create instances of the default interface ITableLog exposed by              
// the CoClass TableLog. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTableLog = class
    class function Create: ITableLog;
    class function CreateRemote(const MachineName: string): ITableLog;
  end;

// *********************************************************************//
// The Class CoDelimLog provides a Create and CreateRemote method to          
// create instances of the default interface IDelimLog exposed by              
// the CoClass DelimLog. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoDelimLog = class
    class function Create: IDelimLog;
    class function CreateRemote(const MachineName: string): IDelimLog;
  end;

// *********************************************************************//
// The Class CoXMLLog provides a Create and CreateRemote method to          
// create instances of the default interface IXMLLog exposed by              
// the CoClass XMLLog. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoXMLLog = class
    class function Create: IXMLLog;
    class function CreateRemote(const MachineName: string): IXMLLog;
  end;

// *********************************************************************//
// The Class CoLogFactory provides a Create and CreateRemote method to          
// create instances of the default interface ILogFactory exposed by              
// the CoClass LogFactory. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoLogFactory = class
    class function Create: ILogFactory;
    class function CreateRemote(const MachineName: string): ILogFactory;
  end;

// *********************************************************************//
// The Class CoUSLogBase provides a Create and CreateRemote method to          
// create instances of the default interface IUSLog exposed by              
// the CoClass USLogBase. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoUSLogBase = class
    class function Create: IUSLog;
    class function CreateRemote(const MachineName: string): IUSLog;
  end;

implementation

uses ComObj;

class function CoFileLog.Create: IFileLog;
begin
  Result := CreateComObject(CLASS_FileLog) as IFileLog;
end;

class function CoFileLog.CreateRemote(const MachineName: string): IFileLog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FileLog) as IFileLog;
end;

class function CoTableLog.Create: ITableLog;
begin
  Result := CreateComObject(CLASS_TableLog) as ITableLog;
end;

class function CoTableLog.CreateRemote(const MachineName: string): ITableLog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TableLog) as ITableLog;
end;

class function CoDelimLog.Create: IDelimLog;
begin
  Result := CreateComObject(CLASS_DelimLog) as IDelimLog;
end;

class function CoDelimLog.CreateRemote(const MachineName: string): IDelimLog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_DelimLog) as IDelimLog;
end;

class function CoXMLLog.Create: IXMLLog;
begin
  Result := CreateComObject(CLASS_XMLLog) as IXMLLog;
end;

class function CoXMLLog.CreateRemote(const MachineName: string): IXMLLog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_XMLLog) as IXMLLog;
end;

class function CoLogFactory.Create: ILogFactory;
begin
  Result := CreateComObject(CLASS_LogFactory) as ILogFactory;
end;

class function CoLogFactory.CreateRemote(const MachineName: string): ILogFactory;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_LogFactory) as ILogFactory;
end;

class function CoUSLogBase.Create: IUSLog;
begin
  Result := CreateComObject(CLASS_USLogBase) as IUSLog;
end;

class function CoUSLogBase.CreateRemote(const MachineName: string): IUSLog;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_USLogBase) as IUSLog;
end;

end.
