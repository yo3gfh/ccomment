/*
    COMMENTPROC, a plugin for Pelle's C IDE
    ---------------------------------------
    Copyright (c) 2002-2020 Adrian Petrila, YO3GFH

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

                                  * * *

    Features
    --------

        - inserts a function commnent header at current cursor/selection
*/

#pragma warn(disable: 2008 2118 2228 2231 2030 2260)

#define         UNICODE

#include        <windows.h>
#include        <commctrl.h>
#include        <addin.h>
#include        <strsafe.h>

#define         IDI_MAIN            8001 // menu icon
#define         IDI_INC             8002 // menu icon
#define         IDD_OPT             1001 // options dlg
#define         IDC_AUTH            4003 // author editbox
#define         IDC_KEYS            4004 // keys cbbox
#define         IDC_CLIP            4005 // copy to clip ckbox
#define         IDC_HELPTXT         4008 // static ctrl with help text

#define         ID_COMMENTPROC      69
#define         ID_FINDPROC         70
#define         ID_SUBCLASS         0xDEAD1969
#define         MAX_LEN             8192
#define         MAX_ALLOCSIZE       0x00100000 // cca 1MB
#define         MAX_LINE            128
#define         MAX_PARAMS          127
#define         MAX_MODIFS          127

typedef 
BOOL ( CALLBACK * CFSEARCHPROC ) 
(
    BOOL valid,    /*a good line?*/
    DWORD lines,
    DWORD crtline,
    WPARAM wParam, 
    LPARAM lParam 
);

typedef
struct tag_cfunction
{
    WCHAR   * params[MAX_PARAMS];   // param. list
    WCHAR   * modifs[MAX_MODIFS];   // return type, modifiers, function name
    WCHAR   * fn_name;              // function name, after StripCrap :-)
    INT_PTR n_par;                  // # of params
    INT_PTR n_mod;                  // # of modifiers
    INT_PTR par_maxw;               // strlen of the longest param.
    INT_PTR mod_maxw;               // strlen of the longest modif.
    BOOL    is_decl;                // is fn. declaration, or definition
    BOOL    is_declspec;            // is there a declspec(...) mod present?
} CFUNCTION;

WPARAM vkeys[7] = { VK_F6, VK_F7, VK_F8, VK_F9, VK_F10, VK_F11, VK_F12 };

HANDLE          g_hMod;             // DLL hmod
HWND            g_hMain;            // main IDE window
WPARAM          g_kbshortcut;       // vkeys[x]
BOOL            g_bcopyclip;        // copy to clippy or not

/*-------------------------- Forward declarations --------------------------*/
static BOOL IsEndl ( WCHAR wc );
static BOOL IsWhite ( WCHAR wc );
static BOOL IsCFunction ( const WCHAR * ws, CFUNCTION * cf );
static BOOL CommentCFunction ( HWND hIde, const WCHAR * csrc );
static BOOL ParseCFunction ( const WCHAR * csrc, CFUNCTION * cf );
static BOOL CopyFunctionInfo ( HWND hIde, const CFUNCTION * cf );
static WCHAR * FILE_Extract_filename ( const WCHAR * src );
static WCHAR * FILE_Extract_path ( const WCHAR * src, BOOL last_bslash );
static WCHAR * FILE_Extract_ext ( const WCHAR * src );

static BOOL SourceCopyFunctionInfo
    ( HWND hDoc, BOOL is_decl, const WCHAR * buf );

static BOOL ReadCCopyFromIni ( const WCHAR * inifile );
static int GetCrtSourceLine ( HWND hDoc );
static HWND GetCrtSourceWnd ( HWND hIde );
static BOOL GetUserSelection ( HWND hIde, WCHAR * wsrc, int cchMax );
static WCHAR * Today ( void );
static WCHAR * IniFromModule ( HMODULE hMod );
static WCHAR * ReadAuthorFromIni ( const WCHAR * inifile );
static WCHAR * StripCrap ( const WCHAR * wsrc, WCHAR * wdest, int cchMax );
static UINT_PTR ReadVKeyFromIni ( const WCHAR * inifile );

static LRESULT CALLBACK DocSubclassProc ( HWND hWnd, UINT msg, WPARAM wParam, 
    LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData );

static INT_PTR CALLBACK OptDlgProc ( HWND hDlg, UINT msg, WPARAM wParam, 
    LPARAM lParam );

static UINT_PTR FindCFDefinitions ( HWND hDoc, 
    CFSEARCHPROC searchproc, LPARAM lParam );
    
static UINT_PTR GetSourceLines ( HWND hDoc, DWORD line, 
    WCHAR * dest, WORD cchMax );

static BOOL CALLBACK FindCDefsProc 
    ( BOOL valid, DWORD lines, DWORD crtline, WPARAM wParam, LPARAM lParam );

/*-@@+@@--------------------------------------------------------------------*/
//       Function: DllMain 
/*--------------------------------------------------------------------------*/
//           Type: BOOL WINAPI 
//    Param.    1: HANDLE hDLL       : 
//    Param.    2: DWORD dwReason    : 
//    Param.    3: LPVOID lpReserved : 
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 27.09.2020
//    DESCRIPTION: entrypoint to our DLL
/*--------------------------------------------------------------------@@-@@-*/
BOOL WINAPI DllMain ( HANDLE hDLL, DWORD dwReason, LPVOID lpReserved )
/*--------------------------------------------------------------------------*/
{
    switch ( dwReason )
    {
        case DLL_PROCESS_ATTACH:
            // Save this, it's not that easy to get hold of it afterwards :-)
            g_hMod = hDLL;
            return TRUE;

        case DLL_PROCESS_DETACH:
            return TRUE;

        default:
            return TRUE;
    }
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: AddInMain 
/*--------------------------------------------------------------------------*/
//           Type: ADDINAPI BOOL WINAPI 
//    Param.    1: HWND hwnd          : main ide window
//    Param.    2: ADDIN_EVENT eEvent : whatever ide just did :-)
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: required export for any dll addin for Pelle ide.
/*--------------------------------------------------------------------@@-@@-*/
ADDINAPI BOOL WINAPI AddInMain ( HWND hwnd, ADDIN_EVENT eEvent )
/*--------------------------------------------------------------------------*/
{
    ADDIN_ADD_COMMAND   add_cmd;
    UINT_PTR            idx;

    switch ( eEvent )
    {
        case AIE_APP_CREATE: // add our menu to "Source" menu
            g_hMain         = hwnd;
            idx             = ReadVKeyFromIni ( IniFromModule ( g_hMod ) );
            g_bcopyclip     = ReadCCopyFromIni ( IniFromModule ( g_hMod ) );
            g_kbshortcut    = vkeys[idx];

            RtlZeroMemory ( &add_cmd, sizeof (ADDIN_ADD_COMMAND));
            add_cmd.cbSize  = sizeof (ADDIN_ADD_COMMAND);
            add_cmd.pszText = L"Add function description";
            add_cmd.hIcon   = LoadImage ( g_hMod, MAKEINTRESOURCE ( IDI_MAIN ),
                IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED );
            add_cmd.id      = ID_COMMENTPROC;
            add_cmd.idMenu  = AIM_MENU_SOURCE;
            AddIn_AddCommand ( hwnd, &add_cmd );

            RtlZeroMemory ( &add_cmd, sizeof (ADDIN_ADD_COMMAND));
            add_cmd.cbSize  = sizeof (ADDIN_ADD_COMMAND);
            add_cmd.pszText = L"Export all functions to header";
            add_cmd.hIcon   = LoadImage ( g_hMod, MAKEINTRESOURCE ( IDI_INC ),
                IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR | LR_SHARED );
            add_cmd.id      = ID_FINDPROC;
            add_cmd.idMenu  = AIM_MENU_SOURCE;
            AddIn_AddCommand ( hwnd, &add_cmd );

            break;

        case AIE_APP_DESTROY:
            g_hMain = NULL;
            AddIn_RemoveCommand ( hwnd, ID_COMMENTPROC );
            AddIn_RemoveCommand ( hwnd, ID_FINDPROC );
            break;

        case AIE_DOC_CREATE:

            if ( AddIn_GetDocumentType ( hwnd ) == AID_SOURCE )
                 // install our wndproc
                SetWindowSubclass ( hwnd, DocSubclassProc, ID_SUBCLASS, 0 );

            break;

        case AIE_DOC_DESTROY:

            if ( AddIn_GetDocumentType ( hwnd ) == AID_SOURCE )
                // remove our subclass
                RemoveWindowSubclass ( hwnd, DocSubclassProc, ID_SUBCLASS ); 

            break;

        default:
            break;
    }

    return TRUE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: AddInCommandExW 
/*--------------------------------------------------------------------------*/
//           Type: ADDINAPI void WINAPI 
//    Param.    1: int idCmd       : menu/toolbar ID that send the command
//    Param.    2: LPCVOID pcvData : custom data
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: event handler for menus/toolbar commands
/*--------------------------------------------------------------------@@-@@-*/
ADDINAPI void WINAPI AddInCommandExW ( int idCmd, LPCVOID pcvData )
/*--------------------------------------------------------------------------*/
{
    WCHAR               wsrc[MAX_LEN];
    WCHAR               temp[MAX_PATH];
    WCHAR               * hdrName, * srcPath, * srcExt;
    ADDIN_DOCUMENT_INFO adi;
    HWND                hHdrDoc, hSrcDoc;

    switch ( idCmd )
    {
        case ID_COMMENTPROC:

            if ( GetUserSelection ( g_hMain, wsrc, ARRAYSIZE(wsrc) ) )
                CommentCFunction ( g_hMain, wsrc );

            break;

        case ID_FINDPROC:
            adi.cbSize = sizeof(ADDIN_DOCUMENT_INFO);

            hSrcDoc = GetCrtSourceWnd ( g_hMain );

            if ( AddIn_GetDocumentInfo ( hSrcDoc, &adi) )
            {
                srcExt = FILE_Extract_ext ( adi.szFilename );

                if ( (srcExt == NULL) || ((lstrcmpiW ( srcExt, L".c") != 0) &&
                    ( lstrcmpiW ( srcExt, L".h" ) != 0)) )
                        break;

                hdrName = FILE_Extract_filename ( adi.szFilename );
                srcPath = FILE_Extract_path ( adi.szFilename, TRUE );

                StringCchPrintfW ( adi.szFilename, 
                    ARRAYSIZE(adi.szFilename), L"%ls%ls.h", 
                    ((srcPath == NULL) ? L"" : srcPath), 
                    ((hdrName == NULL) ? L"untitled" : hdrName));

                hHdrDoc = AddIn_NewDocument ( g_hMain, AID_SOURCE );

                if ( hHdrDoc )
                {
                    AddIn_SetDocumentInfo ( hHdrDoc, &adi );
                    CharUpperW ( hdrName );
                    StringCchPrintfW ( temp, ARRAYSIZE(temp),
                        L"#ifndef _%ls_H\n#define _%ls_H\n\n", 
                        hdrName, hdrName );
                    AddIn_ReplaceSourceSelText ( hHdrDoc, temp );

                    FindCFDefinitions ( hSrcDoc, 
                        (CFSEARCHPROC)FindCDefsProc, (LPARAM)hHdrDoc );

                    StringCchPrintfW ( temp, ARRAYSIZE(temp),
                        L"#endif // _%ls_H\n", hdrName );
                    AddIn_ReplaceSourceSelText ( hHdrDoc, temp );
                }
            }
            else
                MessageBox ( g_hMain, 
                    L"Unable to find current source window :-(", 
                    L"Plugin", MB_OK );
            break;

        default:
            break;
    }
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: AddInSetup 
/*--------------------------------------------------------------------------*/
//           Type: ADDINAPI void WINAPI 
//    Param.    1: HWND hwnd       : IDE window
//    Param.    2: LPCVOID pcvData : custom data
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: exported, called when user clicks the "Options" button from
//                 tools->options menu
/*--------------------------------------------------------------------@@-@@-*/
ADDINAPI void WINAPI AddInSetup ( HWND hwnd, LPCVOID pcvData )
/*--------------------------------------------------------------------------*/
{
    WCHAR   buf[MAX_LEN];
    WCHAR   * ini, * p;

    ini     = IniFromModule ( g_hMod );
    p       = ini;

    if ( ini[0] != L'\0' )
    {
        p += lstrlenW ( ini );

        while ( *p != L'\\' && p != ini )
            p--;

        StringCchPrintfW ( buf, ARRAYSIZE(buf), 
            L"This will create a file named \"%ls\" in /Bin/Addins64/ folder. "
            L"The key \"author=<yourname>\" allows you to customize "
            L"the AUTHOR field from the function description block. "
            L"You can also configure the keyboard shortcut for the menu - "
            L"the key is \"kb_shortcut=<0..6>\", <0> being F6 and <6> F12."
            L"Set \"clip_copy\" key to <1> to also copy fn. description "
            L"to clipboard." , ++p );

        DialogBoxParam ( g_hMod, MAKEINTRESOURCE ( IDD_OPT ), 
            hwnd, OptDlgProc, ( LPARAM ) buf );
    }
}


/*-@@+@@--------------------------------------------------------------------*/
//       Function: GetUserSelection 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: HWND hIde   : main IDE window
//    Param.    2: WCHAR * wsrc: buffer to receive text
//    Param.    3: int cchMax  : wsrc max. size, in characters
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 27.09.2020
//    DESCRIPTION: returns selected text or text from the line under the
//                 cursor. for one line function decl., just place the caret
//                 on the line. for Microsoft style multi-line function
//                 prototypes, you'll have to manually select function
/*--------------------------------------------------------------------@@-@@-*/
static BOOL GetUserSelection ( HWND hIde, WCHAR * wsrc, int cchMax )
/*--------------------------------------------------------------------------*/
{
    WCHAR           temp[256];
    HWND            hDoc;
    ADDIN_RANGE     chr;
    int             line;

    if ( wsrc == NULL || cchMax == 0 ) 
        return FALSE;

    wsrc[0] = L'\0';

    hDoc = AddIn_GetActiveDocument ( hIde );

    if ( hDoc == NULL )
        return FALSE;

    if ( AddIn_GetDocumentType ( hDoc ) == AID_SOURCE )
    {
        if ( !AddIn_GetSourceSel ( hDoc, &chr ) )
            return FALSE;

        if ( chr.iEndPos > chr.iStartPos ) // we have a selection
        {
            if ( (chr.iEndPos - chr.iStartPos) > cchMax )
            {
                StringCchPrintfW ( temp, ARRAYSIZE(temp), 
                    L"Selection is too large. Keep it under %d chars :-)", 
                    cchMax );

                MessageBox ( g_hMain, temp, L"Commentproc Plugin", 
                    MB_OK | MB_ICONINFORMATION );
                return FALSE;
            }

            AddIn_GetSourceSelText ( hDoc, wsrc, cchMax );
        }
        else // just get the line under cursor
        {
            line = AddIn_SourceLineFromChar ( hDoc, chr.iStartPos );
            AddIn_GetSourceLine ( hDoc, line, wsrc, cchMax );
        }
    }

    return TRUE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: IsEndl 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: WCHAR wc : a widechar :-)
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: are we a terminator of lines?
/*--------------------------------------------------------------------@@-@@-*/
static BOOL IsEndl ( WCHAR wc )
/*--------------------------------------------------------------------------*/
{
    return ( wc == L'\r' || wc == L'\n' );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: IsWhite 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: WCHAR wc : 
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: are we whitespace?
/*--------------------------------------------------------------------@@-@@-*/
static BOOL IsWhite ( WCHAR wc )
/*--------------------------------------------------------------------------*/
{
    return ( wc == L' ' || wc == L'\t' );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: IsCFunction 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: const WCHAR * ws: our candidate
//    Param.    2: CFUNCTION * cf  : struct to fill with var. data
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: some crude attempt to assess if the text we got is a 
//                 C function or not :-) - basically, it needs 1 pair of () 
//                 or 2 pairs if declspec is present, also checks if it's 
//                 inside quotes, on a comment line or in n if/else/switch/
//                 ... block and fails accordingly. It can be easily fooled 
//                 and it's not designed to be standalone, but rather as a 
//                 pre-stage for ParseCFunction
/*--------------------------------------------------------------------@@-@@-*/
static BOOL IsCFunction ( const WCHAR * ws, CFUNCTION * cf )
/*--------------------------------------------------------------------------*/
{
    INT_PTR     i, left, right, len;
    BOOL        is_declsp;
    WCHAR       * quot, * comma, 
                * s, * semi_c;
    WCHAR       temp[MAX_LEN];

    left        = 0;
    right       = 0;
    is_declsp   = FALSE;

    quot        = wcschr ( ws, L'"' ); // are there any quotes?

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'=' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'+' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'-' );

    if ( quot != NULL )
        return FALSE;


    quot        = wcschr ( ws, L'?' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'@' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'.' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'!' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'|' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L':' );

    if ( quot != NULL )
        return FALSE;

    quot        = wcschr ( ws, L'&' );

    if ( quot != NULL )
        return FALSE;

    semi_c      = wcschr ( ws, L';' );
    comma       = wcschr ( ws, L',' );
    
    // make a copy for _tcstok
    StringCchCopy ( temp, ARRAYSIZE(temp), ws );

    s           = _wcstok_ms ( temp, L" ()[]{},;" );

    if ( s != NULL )
    {
        if ( ( wcscmp ( s, L"_declspec" )  == 0 )  ||
             ( wcscmp ( s, L"__declspec" ) == 0 ) )
                is_declsp = TRUE;

        if ( ( wcscmp ( s, L"if" )         == 0 )  ||
             ( wcscmp ( s, L"else" )       == 0 )  ||
             ( wcscmp ( s, L"while" )      == 0 )  ||
             ( wcscmp ( s, L"for" )        == 0 )  ||
             ( wcscmp ( s, L"switch" )     == 0 )  ||
             ( wcscmp ( s, L"case" )       == 0 )  ||
             ( wcscmp ( s, L"default" )    == 0 )  ||
             ( wcscmp ( s, L"__except" )   == 0 )  ||
             ( wcscmp ( s, L"defined" )    == 0 )  ||
             ( wcscmp ( s, L"sizeof" )     == 0 )  ||
             ( wcscmp ( s, L"assert" )     == 0 )  ||
             ( wcscmp ( s, L"return" )     == 0 ) )
                return FALSE;

        while ( (s = _wcstok_ms ( NULL,  L" ()[]{},;" )) != NULL )
        {
            if ( ( wcscmp ( s, L"_declspec" )  == 0 )  ||
                 ( wcscmp ( s, L"__declspec" ) == 0 ) )
            {
                is_declsp = TRUE;
                continue;
            }

            if ( ( wcscmp ( s, L"if" )         == 0 )  ||
                 ( wcscmp ( s, L"else" )       == 0 )  ||
                 ( wcscmp ( s, L"while" )      == 0 )  ||
                 ( wcscmp ( s, L"for" )        == 0 )  ||
                 ( wcscmp ( s, L"switch" )     == 0 )  ||
                 ( wcscmp ( s, L"case" )       == 0 )  ||
                 ( wcscmp ( s, L"default" )    == 0 )  ||
                 ( wcscmp ( s, L"__except" )   == 0 )  ||
                 ( wcscmp ( s, L"defined" )    == 0 )  ||
                 ( wcscmp ( s, L"sizeof" )     == 0 )  ||
                 ( wcscmp ( s, L"assert" )     == 0 )  ||
                 ( wcscmp ( s, L"return" )     == 0 ) )
                    return FALSE;
        }
    }

    len     = lstrlenW ( ws );
    i       = len;

    while ( i >= 0)
    {
        if ( ws[i] == L'(' )
        {
            left++;

            if ( semi_c != NULL )
                if ( semi_c < ws+i )
                    return FALSE;

            if ( comma != NULL )
                if ( comma < ws+i )
                    return FALSE; // shouldn't have a comma before '('
        }

        if ( ws[i] == L')' )
        {
            right++;

            if ( semi_c != NULL )
                if ( semi_c < ws+i )
                    return FALSE;

            if ( comma != NULL )
                if ( comma > ws+i )
                    // we've hit some multi-line function call so get out
                    return FALSE;
        }

        i--;
    }

    if ( semi_c != NULL )
        cf->is_decl = ( ws+(len-1) == semi_c ); 
    else
        cf->is_decl = FALSE;

    cf->is_declspec = is_declsp;

    if ( is_declsp )
    {
        if ( (left == 2) && (right == 2) )
            return TRUE;
        else
            return FALSE;
    }

    if ( (left == 1) && (right == 1) )
        return TRUE;

    return FALSE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: StripCrap 
/*--------------------------------------------------------------------------*/
//           Type: static WCHAR * 
//    Param.    1: const WCHAR * wsrc: raw text validated by IsCFunction
//    Param.    2: WCHAR * wdest     : where to put cleaned text
//    Param.    3: int cchMax        : max. destination, in characters
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: strips all CR/LF and extra tabs/spaces from wsrc. null 
//                 terminates wdest
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * StripCrap ( const WCHAR * wsrc, WCHAR * wdest, int cchMax )
/*--------------------------------------------------------------------------*/
{
    INT_PTR i, len, start;
    WCHAR   * p, * s;

    BOOL    skipping        = TRUE;
    BOOL    cr_skipping     = TRUE;
    BOOL    in_comment      = FALSE;
    BOOL    in_macro        = FALSE;
    BOOL    block_comment   = FALSE;

    p       = wdest;
    len     = lstrlenW ( wsrc );
    start   = 0;

    if ( len > cchMax )
        len = cchMax;

    s = wcschr ( wsrc, L';' );

    // see how many ';' we have and skip
    // everything left of them, except the last
    while ( s != NULL )
    {
        s++;
 
        while ( IsWhite(*s) || IsEndl(*s) )
            s++;

        if ( *s == L'\0' )
            break;

        start = (s - wsrc);
        s = wcschr ( wsrc+start, L';' );
    }

    // start where last ';' found
    for ( i = start; i < len; i++ )
    {
        if ( block_comment )
        {
            if ( (i+1 < len) && (wsrc[i] == L'*') && 
                (wsrc[i+1] == L'/'))
            {
                i++;
                block_comment = FALSE;
            }

            continue;
        }

        if ( in_comment )
        {
            if ( IsEndl ( wsrc[i]) )
                in_comment = FALSE;

            continue;
        }

        if ( in_macro )
        {
            if ( IsEndl ( wsrc[i]) )
                in_macro = FALSE;

            continue;
        }

        if ( (i+1 < len) && (wsrc[i] == L'/') && 
            (wsrc[i+1] == '/'))
        {
            i++;
            in_comment = TRUE;
            continue;
        }

        if ( (i+1 < len) && (wsrc[i] == L'/' ) && 
            (wsrc[i+1] == L'*'))
        {
            i++;
            block_comment = TRUE;
            continue;
        }

        if ( wsrc[i] == L'#') 
        {
            in_macro = TRUE;
            continue;
        }

        // try to accomodate the multiline MS style function
        // "beautification" :-)
        // gobble one CR and make it a space
        if ( IsEndl ( wsrc[i]) ) 
        {
            if ( cr_skipping == FALSE )
            {
                *p++ = TEXT(' ');
                cr_skipping = TRUE;
            }
            continue;
        }

        if ( IsWhite(wsrc[i]) )
        {
            if ( skipping == FALSE )
            {
                *p++ = wsrc[i]; // keep at least one tab/space
                skipping = TRUE;
            }
            continue;
        }
        else
        {
            *p++ = wsrc[i];
            skipping = FALSE;
            cr_skipping = FALSE;
        }
    }

    *p = L'\0';
    return wdest;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: CommentCFunction 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: HWND hIde          : main IDE window
//    Param.    2: const WCHAR * csrc : raw text, as returned by 
//                                      GetUserSelection
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: just a wrapper for ParseCFunction and CopyFunctionInfo
/*--------------------------------------------------------------------@@-@@-*/
static BOOL CommentCFunction ( HWND hIde, const WCHAR * csrc )
/*--------------------------------------------------------------------------*/
{
    CFUNCTION   cf;

    if ( !ParseCFunction ( csrc, &cf ) )
        return FALSE;

    return CopyFunctionInfo ( hIde, &cf );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: ParseCFunction 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: const WCHAR * csrc: raw text, as returned by 
//                                     GetUserSelection
//    Param.    2: CFUNCTION * cf    : pointer to a CFUNCTION struct to fill 
//                                     with info
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: This will take a block of text, validate it as a C 
//                 function, then parse it to extract params and modifiers and 
//                 to put them as pointers to null-term. strings in the 
//                 cf->params ans cf->modifs. Will also fill other members of 
//                 CFUNCTION with meaningful info. Returns TRUE if it thinks 
//                 it has a valid C function decl./def. or FALSE otherwise; 
//                 it can return FALSE even if CF_IsCFunction said OK - this 
//                 one checks for a function name and one param at least.
/*--------------------------------------------------------------------@@-@@-*/
static BOOL ParseCFunction ( const WCHAR * csrc, CFUNCTION * cf )
/*--------------------------------------------------------------------------*/
{
    static WCHAR    work_buf[MAX_LEN];
    static WCHAR    next_buf[MAX_LEN];
    static WCHAR    fun_name[MAX_LEN];
    WCHAR           * pos, * p;
    WCHAR           * par_start;
    WCHAR           * mod_start;
    INT_PTR         i, j, len;
    BOOL            is_declsp;

    RtlZeroMemory ( cf, sizeof(CFUNCTION) );
    RtlZeroMemory ( work_buf, sizeof(work_buf) );
    RtlZeroMemory ( next_buf, sizeof(next_buf) );

    // stage 1: clean source from CR/LF, tabs, etc.
    // rough pre-evaluation :-)
    pos             = StripCrap ( csrc, work_buf, ARRAYSIZE(work_buf) );

    if ( !IsCFunction ( pos, cf ) )
        return FALSE;

    // a small hack, clean fn name from any left
    // right curls; could do it in the CF_StripCrap,
    // but the function synopsis won't work correctly.
    // a very good example of bad design :-)
    p               = wcschr ( pos, L'}');

    if ( p != NULL )
    {
        p++;

        while (((*p == L' ')   || 
                (*p == L'}'))  && 
                (*p != L'\0'))
            p++;

        pos = p;
    }

    // save the untouched function
    StringCchCopyW ( fun_name, ARRAYSIZE(fun_name), pos );
    i               = lstrlenW ( fun_name );

    while ( i >= 0 )
    {
        if ( (fun_name[i] == L')') )
            break;

        i--;
    }

    fun_name[i+1]   = L'\0';
    cf->fn_name     = fun_name;

    mod_start       = pos;
    par_start       = pos; // just in case
    i               = 0;
    len             = lstrlenW ( pos );
    is_declsp       = cf->is_declspec;
    
    // stage 2: put a null where '(' is; this will be the boundary between
    // fn. name+modifiers and parameter list; also put a null on ')', since
    // we don't need anything after that
    // if we have declspec(...) replace () with spaces and redo stage 1
    for ( j = 0; j < len; j++ )
    {
        if ( pos[j] == L'(' )
        {
            if ( is_declsp )
                pos[j++] = L' ';
            else
            {
                pos[j++] = L'\0';
                par_start = pos + j;
            }
            continue;
        }
        if ( pos[j] == L')' )
        {
            if ( is_declsp ) // redo stage 1
            {
                pos[j] = L' ';
                is_declsp = FALSE;
                pos = StripCrap ( work_buf, next_buf, ARRAYSIZE(next_buf) );
                mod_start = pos;
                par_start = pos;
                j = 0;
                continue;
            }
            else
                pos[j] = L'\0';
        }
    }
    
    pos             = mod_start;
    len             = lstrlenW ( mod_start );

    // stage 3: parse modifiers and replace every space with a null;
    // set the modifiers pointer table from the CFUNCTION struct
    for ( j = 0; j < len; j++ )
    {
        if ( IsWhite(pos[j]) )
        {
            pos[j++] = L'\0';
            cf->modifs[i++] = mod_start;
            mod_start = pos + j;
        }
        else
        {
            if (pos[j+1] == L'\0')
                cf->modifs[i++] = mod_start;
        }
        // we have a glorious function with 127 modifiers...
        if ( i >= MAX_MODIFS-1 ) 
            break;
    }

    cf->modifs[i]   = 0;
    cf->n_mod       = i;

    if ( i == 0 ) // not ok with 0 modifs.
        return FALSE;

    pos             = par_start;
    i               = 0;
    len             = lstrlenW ( par_start );

    // stage 4: parse parameter list and replace every ',' with a null;
    // set the params pointer table from the CFUNCTION struct
    for ( j = 0; j < len; j++ )
    {
        if (pos[j] == L',')
        {
            pos[j++] = L'\0';
            // skip leading space, if any
            while ( IsWhite(*par_start) && *par_start != L'\0' ) 
                par_start++;

            cf->params[i++] = par_start;
            par_start = pos + j;
        }
        else
        {
            if (pos[j+1] == L'\0')
            {
                while ( IsWhite(*par_start) && *par_start != L'\0' )
                    par_start++;

                cf->params[i++] = par_start;
            }
        }
        // ...or another glorious function with 127 params :-)
        if ( i >= MAX_PARAMS-1 ) 
            break;
    }
    
    cf->params[i]   = 0; 
    cf->n_par       = i;

    if ( i == 0 )
        return FALSE; // not ok with 0 params

    // save the longest param/mod for later 
    for ( j = 0; j < cf->n_mod; j++ )
    {
        len = lstrlenW ( cf->modifs[j] );

        if ( cf->mod_maxw < len )
            cf->mod_maxw = len;
    }

    for ( j = 0; j < cf->n_par; j++ )
    {
        len = lstrlenW ( cf->params[j] );

        if ( cf->par_maxw < len )
            cf->par_maxw = len;
    }

    return TRUE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: CopyFunctionInfo 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: HWND hIde            : main IDE window
//    Param.    2: const CFUNCTION * cf : addr. of a CFUNCTION struct, 
//                                        filled by ParseCFunction 
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: take a CFUNCTION *, allocate mem. , make it into nice 
//                 formatted text and shove it around the current function 
//                 declaration/definition. if g_bcopyclip is set, also copy 
//                 it to clipboard. NOTE: don't GlobalFree mem. that you give 
//                 to Clippy :-)
/*--------------------------------------------------------------------@@-@@-*/
static BOOL CopyFunctionInfo ( HWND hIde, const CFUNCTION * cf )
/*--------------------------------------------------------------------------*/
{
    WCHAR       * ini, * buf, line[MAX_LINE*2];
    HGLOBAL     hMem;
    INT_PTR     buf_size, alloc_size;
    INT_PTR     i;
    HWND        hDoc;

    if ( cf == NULL )
        return FALSE;

    if ( cf->n_mod == 0 || cf->n_par == 0 )
        return FALSE;

    // n_mod lines for modifiers, n_par lines 
    // for param list, 3 lines header+footer, 
    // 3 for author, date, and description - 128 chars each
    buf_size    = ((cf->n_mod + cf->n_par + 6) * MAX_LINE);
    alloc_size  = buf_size * sizeof(WCHAR); 
    
    if ( alloc_size > MAX_ALLOCSIZE )
        return FALSE;

    hMem = GlobalAlloc ( GHND, alloc_size );

    if ( hMem == NULL )
        return FALSE;

    ini         = IniFromModule ( g_hMod );
    buf         = (WCHAR *)GlobalLock ( hMem );

    StringCchPrintfW ( buf, buf_size,
        L"/*-@@+@@-----------------------------------"
        L"---------------------------------*/\n" );
    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"//       Function: %ls \n", cf->modifs[cf->n_mod-1] );
    StringCchCatW ( buf, buf_size, line );
    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"/*-----------------------------------------"
        L"---------------------------------*/\n" );
    StringCchCatW ( buf, buf_size, line );
    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"//           Type: " );
    StringCchCatW ( buf, buf_size, line );

    for ( i = 0; i < cf->n_mod-1; i++ )
    {
        StringCchPrintfW ( line, ARRAYSIZE(line), L"%ls ", cf->modifs[i] );
        StringCchCatW ( buf, buf_size, line );
    }

    StringCchPrintfW ( line, ARRAYSIZE(line),     L"\n" );
    StringCchCatW ( buf, buf_size, line );

    for ( i = 0; i < cf->n_par; i++ )
    {
        StringCchPrintfW ( line, ARRAYSIZE(line),
            L"//    Param. %4zd: %-*ls: \n", i+1, 
            (int)cf->par_maxw, cf->params[i] );

        StringCchCatW ( buf, buf_size, line );
    }

    StringCchPrintfW ( line, ARRAYSIZE(line), 
        L"/*-------------------------------------------"
        L"-------------------------------*/\n" );
    StringCchCatW ( buf, buf_size, line );
    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"//         AUTHOR: %ls\n//           DATE: %ls"
        L"\n//    DESCRIPTION: <lol>\n//\n", 
        ReadAuthorFromIni ( ini ), Today() );

    StringCchCatW ( buf, buf_size, line );
    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"/*--------------------------------------------"
        L"------------------------@@-@@-*/\n" );
    StringCchCatW ( buf, buf_size, line );

    hDoc = GetCrtSourceWnd ( hIde );

    // put info in the source as well
    if ( hDoc != NULL )
        SourceCopyFunctionInfo ( hDoc, cf->is_decl, buf );

    GlobalUnlock ( hMem );

    if ( !g_bcopyclip )
    {
        GlobalFree ( hMem );
        return TRUE; // nothing more to do, so close the shop
    }

    if ( !OpenClipboard ( g_hMain ) )
    {
        GlobalFree ( hMem );
        return FALSE; //all this work in vain lol
    }

    EmptyClipboard();

    // give hMem to clipboard, don't GlobalFree :-)
    if ( !SetClipboardData ( CF_UNICODETEXT, hMem ) ) 
    {
        GlobalFree ( hMem );
        CloseClipboard();
        return FALSE;
    }

    CloseClipboard();

    StringCchPrintfW ( line, ARRAYSIZE(line),
        L"Function \"%ls\" description copied to clipboard.", 
        cf->modifs[cf->n_mod-1] );

    AddIn_WriteOutputW ( hIde, line );

    return TRUE;
}


/*-@@+@@--------------------------------------------------------------------*/
//       Function: Today 
/*--------------------------------------------------------------------------*/
//           Type: static WCHAR * 
//    Param.    1: void : 
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: return current date in string format (dd.mm.yyyy)
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * Today ( void )
/*--------------------------------------------------------------------------*/
{
    static WCHAR    buf[128];
    SYSTEMTIME      st;

    buf[0] = '\0';
    GetLocalTime ( &st );
    StringCchPrintfW ( buf, ARRAYSIZE(buf), 
        L"%02u.%02u.%u", st.wDay, st.wMonth, st.wYear );

    return buf;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: SourceCopyFunctionInfo 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: HWND hDoc         : handle to source window
//    Param.    2: BOOL is_decl      : it's a function declaration?
//    Param.    3: const WCHAR * buf : formatted text
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: insert function info above function definition/declaration;
//                 if we have a definition, also insert a comment bar just 
//                 above the opening '{'
/*--------------------------------------------------------------------@@-@@-*/
static BOOL SourceCopyFunctionInfo ( HWND hDoc, BOOL is_decl, 
    const WCHAR * buf )
/*--------------------------------------------------------------------------*/
{
    ADDIN_RANGE     ar;
    INT_PTR         line, retries, lcount;
    WCHAR           temp[MAX_LEN];

    if ( buf == NULL )
        return FALSE;

    if ( *buf == L'\0' )
        return FALSE;

    // get current line and go to line start
    line = GetCrtSourceLine ( hDoc );

    if ( line > 0 )
    {
        AddIn_GetSourceLine ( hDoc, line-1, temp, ARRAYSIZE(temp) );

        if ( wcsstr ( temp, L"@@-@@" ) != NULL )
            return FALSE;
    }

    ar.iEndPos = ar.iStartPos = AddIn_SourceLineIndex ( hDoc, line );

    // move caret and insert formatted text
    AddIn_SetSourceSel ( hDoc, &ar );
    AddIn_ReplaceSourceSelText ( hDoc, buf );

    // no use in adding body separator since 
    // we have just a function declaration
    // so we get out :-)
    if ( is_decl )
        return TRUE;

    // update line idx, seek for first curly for a max. 
    // of 30 lines (or file end,
    // whichever comes first
    line = GetCrtSourceLine ( hDoc );
    line++;
    retries = 0;
    lcount = AddIn_GetSourceLineCount ( hDoc );

    do
    {
        if ( AddIn_GetSourceLine ( hDoc, line, temp, 
            ARRAYSIZE(temp) ) == -1 )
                break;

        if ( wcschr ( temp, L'{' ) != NULL )
        {
            ar.iEndPos = ar.iStartPos = AddIn_SourceLineIndex ( hDoc, line );
            AddIn_SetSourceSel ( hDoc, &ar );
            AddIn_ReplaceSourceSelText 
                ( hDoc, L"/*---------------------------------------"
                    L"-----------------------------------*/\n" );
            break;
        }

        line++;
        retries++;
    } 
    while ( retries < 31 && line < lcount );

    return TRUE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: GetCrtSourceLine 
/*--------------------------------------------------------------------------*/
//           Type: static int 
//    Param.    1: HWND hDoc : handle to current source window
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: returns the source line which contains the caret
/*--------------------------------------------------------------------@@-@@-*/
static int GetCrtSourceLine ( HWND hDoc )
/*--------------------------------------------------------------------------*/
{
    ADDIN_RANGE     chr;
    int             line;

    if ( AddIn_GetDocumentType ( hDoc ) == AID_SOURCE )
    {
        if ( !AddIn_GetSourceSel ( hDoc, &chr ) )
            return -1;

        line = AddIn_SourceLineFromChar ( hDoc, chr.iStartPos );
        return line;
    }

    return -1;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: GetCrtSourceWnd 
/*--------------------------------------------------------------------------*/
//           Type: static HWND 
//    Param.    1: HWND hIde : handle to main IDE window
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: returns current source window handle
/*--------------------------------------------------------------------@@-@@-*/
static HWND GetCrtSourceWnd ( HWND hIde )
/*--------------------------------------------------------------------------*/
{
    HWND    hDoc;

    hDoc = AddIn_GetActiveDocument ( hIde );

    if ( hDoc != NULL )
    {
        if ( AddIn_GetDocumentType ( hDoc ) == AID_SOURCE )
            return hDoc;
    }

    return NULL;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: IniFromModule 
/*--------------------------------------------------------------------------*/
//           Type: static WCHAR * 
//    Param.    1: HMODULE hMod : our DLL hmodule
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: create a full path ini name from the module name loaded 
//                 as hMod; we assume that the module fname contains a '.'
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * IniFromModule ( HMODULE hMod )
/*--------------------------------------------------------------------------*/
{
    static WCHAR    buf[MAX_PATH];
    WCHAR           * p;

    buf[0]  = L'\0';
    p       = buf;

    if ( GetModuleFileNameW ( hMod, buf, ARRAYSIZE(buf) ) )
    {
        while ( *p != L'.' && *p != L'\0' )
            p++;
        
        if ( *p == L'\0' )
            return buf;

        *(p+1) = L'i';
        *(p+2) = L'n';
        *(p+3) = L'i';
        *(p+4) = L'\0';
    }
    
    return buf;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: ReadAuthorFromIni 
/*--------------------------------------------------------------------------*/
//           Type: static WCHAR * 
//    Param.    1: const WCHAR * inifile : full path to config file
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: read author name from <plugindll.ini> from the same folder 
//                 as the addin (Addins64)
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * ReadAuthorFromIni ( const WCHAR * inifile )
/*--------------------------------------------------------------------------*/
{
    static WCHAR buf[MAX_PATH];

    GetPrivateProfileStringW ( L"settings", L"author", L"<bugmeister>", buf, 
        ARRAYSIZE(buf), inifile );

    return buf;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: ReadVKeyFromIni 
/*--------------------------------------------------------------------------*/
//           Type: static UINT_PTR 
//    Param.    1: const WCHAR * inifile : full path to config file
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: read menu kb. shortcut from ini (index in the vkeys table))
/*--------------------------------------------------------------------@@-@@-*/
static UINT_PTR ReadVKeyFromIni ( const WCHAR * inifile )
/*--------------------------------------------------------------------------*/
{
    return (UINT_PTR) GetPrivateProfileIntW ( L"settings", L"kb_shortcut", 6, 
        inifile ); // return F12 if no config
}


/*-@@+@@--------------------------------------------------------------------*/
//       Function: ReadCCopyFromIni 
/*--------------------------------------------------------------------------*/
//           Type: static BOOL 
//    Param.    1: const WCHAR * inifile : full path to config file
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: read g_bcopyclip value from ini
/*--------------------------------------------------------------------@@-@@-*/
static BOOL ReadCCopyFromIni ( const WCHAR * inifile )
/*--------------------------------------------------------------------------*/
{
    return (BOOL) GetPrivateProfileIntW ( L"settings", L"clip_copy", 0, 
        inifile );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: DocSubclassProc 
/*--------------------------------------------------------------------------*/
//           Type: static LRESULT CALLBACK 
//    Param.    1: HWND hWnd           : subclassed window handle
//    Param.    2: UINT msg            : message to process
//    Param.    3: WPARAM wParam       : message wParam
//    Param.    4: LPARAM lParam       : message lParam
//    Param.    5: UINT_PTR uIdSubclass: unique subclass ID
//    Param.    6: DWORD_PTR dwRefData : custom data (unused)
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: this is the new wndproc for the IDE source window after
//                 being subclassed. basically, we check for pressed key 
//                 against whatever shortcut from config file.
/*--------------------------------------------------------------------@@-@@-*/
static LRESULT CALLBACK DocSubclassProc ( HWND hWnd, UINT msg, WPARAM wParam, 
    LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData )
/*--------------------------------------------------------------------------*/
{
    if ( msg == WM_KEYDOWN )
    {
        if ( wParam == g_kbshortcut )
        {
            AddInCommandExW ( ID_COMMENTPROC, (LPCVOID)uIdSubclass );
            return 0;
        }
    }

    return DefSubclassProc ( hWnd, msg, wParam, lParam );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: OptDlgProc 
/*--------------------------------------------------------------------------*/
//           Type: static INT_PTR CALLBACK 
//    Param.    1: HWND hDlg     : 
//    Param.    2: UINT msg      : 
//    Param.    3: WPARAM wParam : 
//    Param.    4: LPARAM lParam : on WM_INITDIALOG, lParam contains a pointer
//                 to the helptext
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 03.10.2020
//    DESCRIPTION: dialog procedure for the options dialog.
/*--------------------------------------------------------------------@@-@@-*/
static INT_PTR CALLBACK OptDlgProc ( HWND hDlg, UINT msg, WPARAM wParam, 
    LPARAM lParam )
/*--------------------------------------------------------------------------*/
{
    WCHAR * keys[] = {L"F6", L"F7", L"F8", L"F9", L"F10", L"F11", L"F12"};
    WCHAR       * ini;
    WCHAR       temp[256];
    HWND        hKlist, hEauth;
    UINT_PTR     i;
    DWORD       ckstates[2] = { BST_UNCHECKED, BST_CHECKED };
    BOOL        checked;


    switch ( msg )
    {
        case WM_INITDIALOG:
            ini     = IniFromModule ( g_hMod );
            hKlist  = GetDlgItem ( hDlg, IDC_KEYS );
            hEauth  = GetDlgItem ( hDlg, IDC_AUTH );

            for ( i = 0; i < 7; i++ )
                SendMessage ( hKlist, CB_ADDSTRING, 0, 
                    ( LPARAM )( LPCTSTR )keys[i] );

            SendMessage ( hKlist, CB_SETCURSEL, ( WPARAM )
                (ReadVKeyFromIni ( ini )), 0 );

            i = ReadCCopyFromIni ( ini );
            CheckDlgButton ( hDlg, IDC_CLIP, ckstates[i] );
             // read author from ini
            SetDlgItemTextW ( hDlg, IDC_AUTH, ReadAuthorFromIni ( ini ) );
            // set help text
            SetDlgItemTextW ( hDlg, IDC_HELPTXT, (LPCWSTR)lParam ); 
            SetFocus ( hEauth );
            break;
        
        case WM_CLOSE:
            EndDialog ( hDlg, 0 );
            break;

        case WM_COMMAND:
            if ( lParam )
            {
                if ( HIWORD ( wParam ) == BN_CLICKED )
                {
                    switch ( LOWORD ( wParam ) )
                    {
                        case IDCANCEL:
                            SendMessage ( hDlg, WM_CLOSE, 0, 0 );
                            break;

                        case IDOK:
                            ini     = IniFromModule ( g_hMod );
                            hEauth  = GetDlgItem ( hDlg, IDC_AUTH );
                            checked = ( IsDlgButtonChecked 
                                ( hDlg, IDC_CLIP ) == BST_CHECKED );
                            g_bcopyclip = checked;
                            StringCchPrintfW ( temp, ARRAYSIZE(temp), 
                                L"%u", checked );
                            WritePrivateProfileStringW ( L"settings", 
                                L"clip_copy", temp, ini );

                            if ( GetWindowTextLengthW ( hEauth ) != 0 )
                            {
                                GetDlgItemTextW ( hDlg, IDC_AUTH, temp, 
                                    ARRAYSIZE(temp) );
                                WritePrivateProfileStringW ( L"settings", 
                                    L"author", temp, ini );
                            }

                            i = SendDlgItemMessage ( hDlg , IDC_KEYS, 
                                CB_GETCURSEL, (WPARAM) 0, (LPARAM) 0 );

                            if ( i >= 0 )
                            {
                                StringCchPrintfW ( temp, ARRAYSIZE(temp), 
                                    L"%zu", i );
                                WritePrivateProfileStringW ( L"settings", 
                                    L"kb_shortcut", temp, ini );
                                g_kbshortcut = vkeys[i];
                            }

                            EndDialog ( hDlg, IDOK );
                            break;
                    }
                }
            }

        default:
            return FALSE;
            break;
    }

    return TRUE;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: FILE_Extract_filename
/*--------------------------------------------------------------------------*/
//           Type: WCHAR *
//    Param.    1: const WCHAR * src : full path
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 28.09.2020
//    DESCRIPTION: Extract filename, no ext.
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * FILE_Extract_filename ( const WCHAR * src )
/*--------------------------------------------------------------------------*/
{
    DWORD           idx, len, end;
    static WCHAR    temp[MAX_PATH+1];

    if ( src == NULL )
        return NULL;

    len             = lstrlenW ( src );
    idx             = len-1;
    end             = 0;

    while ( ( src[idx] != L'\\' ) && ( idx != 0 ) )
    {
        if ( src[idx] == L'.' )
            end = idx;

        idx--;
    }

    if ( ( idx == 0 ) || ( ( end-idx ) >= MAX_PATH ) )
        return NULL;

    lstrcpynW ( temp, src+idx+1, end-idx );

    return temp;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: FILE_Extract_path
/*--------------------------------------------------------------------------*/
//           Type: WCHAR *
//    Param.    1: const WCHAR * src: a full path
//    Param.    2: BOOL last_bslash : wether to add the last '\' or not
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 28.09.2020
//    DESCRIPTION: Returns a filepath, with or without the last backslash
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * FILE_Extract_path ( const WCHAR * src, BOOL last_bslash )
/*--------------------------------------------------------------------------*/
{
    DWORD           idx;
    static WCHAR    temp[MAX_PATH+1];

    if ( src == NULL )
        return NULL;

    idx = lstrlenW ( src )-1;

    if ( idx >= MAX_PATH )
        return NULL;

    while ( ( src[idx] != L'\\') && ( idx != 0 ) )
        idx--;

    if ( idx == 0 )
        return NULL;

    if ( last_bslash )
        idx++;

    lstrcpynW ( temp, src, idx+1 );

    return temp;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: FILE_Extract_ext
/*--------------------------------------------------------------------------*/
//           Type: WCHAR *
//    Param.    1: const WCHAR * src : hopefully, a good filename :)
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 28.09.2020
//    DESCRIPTION: Returns a pointer to the file extension in src,
//                 or NULL on err.
/*--------------------------------------------------------------------@@-@@-*/
static WCHAR * FILE_Extract_ext ( const WCHAR * src )
/*--------------------------------------------------------------------------*/
{
    DWORD   idx;

    if ( src == NULL )
        return NULL;

    idx = lstrlenW ( src )-1;

    while ( ( src[idx] != L'.' ) && ( idx != 0 ) )
        idx--;

    if ( idx == 0 )
        return NULL;

    return (WCHAR *)( src+idx );
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: GetSourceLines 
/*--------------------------------------------------------------------------*/
//           Type: static INT_PTR 
//    Param.    1: HWND hDoc   : source doc
//    Param.    2: DWORD line  : crt line
//    Param.    3: WCHAR * dest: buffer to receive text
//    Param.    4: WORD cchMax : max buf len, in chars
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 09.10.2020
//    DESCRIPTION: Go through the lines of a C source doc until a function
//                 definition is found and copy it to dest. Returns the #
//                 of lines consumed.
/*--------------------------------------------------------------------@@-@@-*/
static UINT_PTR GetSourceLines ( HWND hDoc, DWORD line, 
    WCHAR * dest, WORD cchMax )
/*--------------------------------------------------------------------------*/
{
    INT_PTR     i;
    UINT_PTR    lines, len, j, 
                left_curl, right_curl, 
                right_par;
    WCHAR       buf[MAX_LEN];
    WCHAR       * comm_start, * comm_end, 
                * lcomm, * p;
    BOOL        in_quote, in_comm;

    if ( dest == NULL || cchMax == 0 )
        return 0;

    lines       = 0;
    len         = 0;
    left_curl   = 0;
    right_curl  = 0;
    right_par   = 0;
    dest[0]     = L'\0';
    in_comm     = FALSE;

    while ( ((WORD)len) < cchMax )
    {
        in_quote = FALSE;

        if ( AddIn_GetSourceLine ( hDoc, (int)(line+lines), 
                buf, ARRAYSIZE(buf)) == -1 )
        {
            lines++;
            break;
        }

        comm_start = wcsstr ( buf, L"/*");
        comm_end = wcsstr ( buf, L"*/");
        lcomm = wcsstr ( buf, L"//");
        p = buf;

        // skip starting whitespace
        while ( IsWhite ( *p ) && *p != L'\0' )
            p++;

        // weak hack, lol
        if ( wcsstr ( p, L"else" ) != NULL )
        {
            lines++;
            continue;
        }

        if ( comm_start )
        {
            if ( comm_end == NULL )
            {
                in_comm = TRUE;
                lines++;
                continue;
            }
        }

        if ( comm_end )
        {
            if ( comm_start == NULL )
            {
                in_comm = FALSE;
                lines++;
                continue;
            }
        }

        if ( lcomm == p || in_comm == TRUE )
        {
            lines++;
            continue;
        }

        i = lstrlenW ( p );
        j = i;

        while ( i >= 0 )
        {
            if ( p[i] == L'\'' || p[i] == L'\"' )
                in_quote = !in_quote;

            if ( !in_quote )
            {
                if ( (p[i] == L')') && ((p+i > comm_end) || 
                    (p+i < comm_start) || (p+i < lcomm)) )
                        right_par++;

                if ( (p[i] == L'{') && ((p+i > comm_end) || 
                    (p+i < comm_start) || (p+i < lcomm)) )
                        left_curl++;

                if ( (p[i] == L'}') && ((p+i > comm_end) || 
                    (p+i < comm_start) || (p+i < lcomm)) )
                        right_curl++;
            }

            i--;
        }

        if ( left_curl > 0 )
        {
            if ( (left_curl != right_curl) && (right_par == 0) )
            {
                lines++;
                continue;
            }
        }

        StringCchCatW ( dest, cchMax-((WORD)len), p );
        len += j;
        lines++;

        if ( left_curl > 0 )
            break;
    }

    return lines-1;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: FindCFDefinitions 
/*--------------------------------------------------------------------------*/
//           Type: INT_PTR 
//    Param.    1: HWND hDoc              : source doc
//    Param.    2: CFSEARCHPROC searchproc: callback to exec when a def.
//                                          is found 
//    Param.    3: LPARAM lParam          : user def for searchproc
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 09.10.2020
//    DESCRIPTION: Search for function definitions in a source doc by calling
//                 GetSourceLines. Call searchproc for every hit. Returns
//                 total number of chars sent to searchproc. 
/*--------------------------------------------------------------------@@-@@-*/
static UINT_PTR FindCFDefinitions ( HWND hDoc, 
    CFSEARCHPROC searchproc, LPARAM lParam )
/*--------------------------------------------------------------------------*/
{
    UINT_PTR    i, len, lines;
    UINT_PTR    size;
    WCHAR       line[MAX_LEN];
    CFUNCTION   cf;
    BOOL        in_comment;

    lines       = AddIn_GetSourceLineCount ( hDoc );

    size        = 0;
    len         = 0;

    in_comment  = FALSE;

    for ( i = 0; i < lines; i++ )
    {
        i += GetSourceLines ( hDoc, (DWORD)i, line, ARRAYSIZE(line) );

        if ( ParseCFunction ( line, &cf ) )
        {
            if ( cf.is_decl )
                continue;

            // can't be a def. with function name only (or no function name)
            if ( cf.n_mod == 1 ) 
                continue;

            StringCchCopyW ( line, ARRAYSIZE(line), cf.fn_name );
            len = lstrlenW (cf.fn_name);
            StringCchCatW ( line, ARRAYSIZE(line), L";" );
            
            if (len < (ARRAYSIZE(line)-1))
                len += 1;

            if ( !searchproc ( TRUE, (DWORD)lines, 
                    (DWORD)(i+1), (WPARAM)line, lParam ) )
            {
                size += len;
                break;
            }

            size += len;
        }
    }

    return size;
}

/*-@@+@@--------------------------------------------------------------------*/
//       Function: FindCDefsProc 
/*--------------------------------------------------------------------------*/
//           Type: BOOL CALLBACK 
//    Param.    1: BOOL valid    : is valid? 
//    Param.    2: DWORD lines   : how many lines
//    Param.    3: DWORD crtline : current line
//    Param.    4: WPARAM wParam : function name
//    Param.    5: LPARAM lParam : custom data
/*--------------------------------------------------------------------------*/
//         AUTHOR: Adrian Petrila, YO3GFH
//           DATE: 09.10.2020
//    DESCRIPTION: Callback for FindCFDefinitions, called each time a fn.
//                 definition is found. 
/*--------------------------------------------------------------------@@-@@-*/
static BOOL CALLBACK FindCDefsProc ( BOOL valid, DWORD lines, 
    DWORD crtline, WPARAM wParam, LPARAM lParam )
/*--------------------------------------------------------------------------*/
{
    WCHAR   temp[2048];

    if ( valid )
    {
        StringCchPrintf ( temp, ARRAYSIZE(temp), L"%ls // line %lu\n", 
            (WCHAR *)wParam, crtline );

        AddIn_ReplaceSourceSelText ( (HWND)lParam, temp );
    }

    return TRUE;
}
