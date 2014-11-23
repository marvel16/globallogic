/****************************** Module Header ******************************\
Module Name:  FileContextMenuExt.cpp
Project:      CppShellExtContextMenuHandler
Copyright (c) Microsoft Corporation.

The code sample demonstrates creating a Shell context menu handler with C++. 

A context menu handler is a shell extension handler that adds commands to an 
existing context menu. Context menu handlers are associated with a particular 
file class and are called any time a context menu is displayed for a member 
of the class. While you can add items to a file class context menu with the 
registry, the items will be the same for all members of the class. By 
implementing and registering such a handler, you can dynamically add items to 
an object's context menu, customized for the particular object.

The example context menu handler adds the menu item "Display File Name (C++)"
to the context menu when you right-click a .cpp file in the Windows Explorer. 
Clicking the menu item brings up a message box that displays the full path 
of the .cpp file.

This source is subject to the Microsoft Public License.
See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
All other rights reserved.

THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/

#include "FileContextMenuExt.h"
#include "resource.h"
#include <strsafe.h>
#include <Shlwapi.h>
#include <string>
#include <vector>
#include <fstream>
#include <set>

#pragma comment(lib, "shlwapi.lib")


extern HINSTANCE g_hInst;
extern long g_cDllRef;

#define IDM_DISPLAY             0  // The command's identifier offset

void WriteToLogFile(const std::multiset<std::wstring> &wstr)
{
	using namespace std;
	wfstream logFile;
	wstring tempStr;
	multiset<wstring> tempSet;
	
	tempSet=wstr;
	
	logFile.open(L"C:\\Log\\logFile.txt",ios::in);
	while(getline(logFile, tempStr))
	{
		if(tempStr!=L"\n")
			tempSet.insert(tempStr);
	}
	logFile.close();
	
	logFile.open(L"C:\\Log\\logFile.txt", ios::out);
	for(multiset<wstring>::iterator it = tempSet.begin(); it != tempSet.end(); it++)
		logFile << *it << endl; //writing into file string
	
}

FileContextMenuExt::FileContextMenuExt(void) : m_cRef(1), 
    m_pszMenuText(L"&Display File Name (C++)"),
    m_pszVerb("cppdisplay"),
    m_pwszVerb(L"cppdisplay"),
    m_pszVerbCanonicalName("CppDisplayFileName"),
    m_pwszVerbCanonicalName(L"CppDisplayFileName"),
    m_pszVerbHelpText("Display File Name (C++)"),
    m_pwszVerbHelpText(L"Display File Name (C++)")
{
    InterlockedIncrement(&g_cDllRef);

    // Load the bitmap for the menu item. 
    // If you want the menu item bitmap to be transparent, the color depth of 
    // the bitmap must not be greater than 8bpp.
    m_hMenuBmp = LoadImage(g_hInst, MAKEINTRESOURCE(IDB_OK), 
        IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE | LR_LOADTRANSPARENT);
}

FileContextMenuExt::~FileContextMenuExt(void)
{
    if (m_hMenuBmp)
    {
        DeleteObject(m_hMenuBmp);
        m_hMenuBmp = NULL;
    }

    InterlockedDecrement(&g_cDllRef);
}


void FileContextMenuExt::OnVerbDisplayFileName(HWND hWnd)
{
    wchar_t szMessage[300];
	
    if (SUCCEEDED(StringCchPrintf(szMessage, ARRAYSIZE(szMessage), 
        L"The selected file is:\r\n\r\n%s", this->m_szSelectedFile)))
    {
        MessageBox(hWnd, szMessage, L"CppShellExtContextMenuHandler", MB_OK);
	}
}


#pragma region IUnknown

// Query to the interface the component supported.
IFACEMETHODIMP FileContextMenuExt::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(FileContextMenuExt, IContextMenu),
        QITABENT(FileContextMenuExt, IShellExtInit), 
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

// Increase the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

// Decrease the reference count for an interface on an object.
IFACEMETHODIMP_(ULONG) FileContextMenuExt::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }

    return cRef;
}

#pragma endregion


#pragma region IShellExtInit

// Initialize the context menu handler.
IFACEMETHODIMP FileContextMenuExt::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (NULL == pDataObj)
    {
        return E_INVALIDARG;
    }

    HRESULT hr = E_FAIL;

    FORMATETC fe = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
    STGMEDIUM stm;

    // The pDataObj pointer contains the objects being acted upon. In this 
    // example, we get an HDROP handle for enumerating the selected files and 
    // folders.
    if (SUCCEEDED(pDataObj->GetData(&fe, &stm)))
    {
        // Get an HDROP handle.
        HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
        if (hDrop != NULL)
        {
            // Determine how many files are involved in this operation. This 
            // code sample displays the custom context menu item when few files are selected. 
            UINT nFiles = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
            
            for (UINT i = 0; i < nFiles;i++)
			{
				// Get the path of the file.
				if (0 == DragQueryFileW(hDrop, i, m_szSelectedFile,
					ARRAYSIZE(m_szSelectedFile)))
				{
					continue;
				}
				fileName.insert(m_szSelectedFile);

				hr = S_OK;
			}
			

            GlobalUnlock(stm.hGlobal);
        }

        ReleaseStgMedium(&stm);
    }

    // If any value other than S_OK is returned from the method, the context 
    // menu item is not displayed.
    return hr;
}

#pragma endregion


#pragma region IContextMenu

//
//   FUNCTION: FileContextMenuExt::QueryContextMenu
//
//   PURPOSE: The Shell calls IContextMenu::QueryContextMenu to allow the 
//            context menu handler to add its menu items to the menu. It 
//            passes in the HMENU handle in the hmenu parameter. The 
//            indexMenu parameter is set to the index to be used for the 
//            first menu item that is to be added.
//
IFACEMETHODIMP FileContextMenuExt::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    // If uFlags include CMF_DEFAULTONLY then we should not do anything.
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    // Use either InsertMenu or InsertMenuItem to add menu items.
    // Learn how to add sub-menu from:
   


    MENUITEMINFO mii = { sizeof(mii) };
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_DISPLAY;
    mii.fType = MFT_STRING;
    mii.dwTypeData = m_pszMenuText;
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = static_cast<HBITMAP>(m_hMenuBmp);
    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add a separator.
    MENUITEMINFO sep = { sizeof(sep) };
    sep.fMask = MIIM_TYPE;
    sep.fType = MFT_SEPARATOR;
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &sep))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Return an HRESULT value with the severity set to SEVERITY_SUCCESS. 
    // Set the code value to the offset of the largest command identifier 
    // that was assigned, plus one (1).
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(IDM_DISPLAY + 1));
}


//
//   FUNCTION: FileContextMenuExt::InvokeCommand
//
//   PURPOSE: This method is called when a user clicks a menu item to tell 
//            the handler to run the associated command. The lpcmi parameter 
//            points to a structure that contains the needed information.
//
IFACEMETHODIMP FileContextMenuExt::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	typedef std::pair<unsigned int, std::wstring> PAIR;
	std::multiset <std::wstring> fileName_info;
	std::vector <PAIR> vec; //vector of pairs string-byte for as return value from threadpool
	for (auto it = fileName.begin(); it != fileName.end(); ++it)
		vec.push_back(std::make_pair(0, *it));
		
	std::vector<std::future<PAIR>> results;
	ThreadPool pool(4); //creating threadpool with 4 threads
	//do tasks (count per byte summ of droped files
	for (int i = 0; i < vec.size(); ++i)
	{
		results.push_back(pool.enqueue(
			[=]()-> std::pair<unsigned int, std::wstring>
		{
			PAIR pairres(vec[i]);
			std::fstream file(pairres.second, std::ios_base::in | std::ios_base::binary);
			unsigned int sum = 0;
			char buff;

			if (!file.is_open())
			{
				//std::cerr << "\nFile wasnt opened...\n";
				throw std::runtime_error("(CheckSum::hashSum : File wasnt opened...");
				
			}

			while (file.read(&buff, 1))
				sum += (unsigned int)(buff);
			pairres.first = sum;
			return pairres;
		}
		));
	}
	//
	for (auto it = results.begin(); it != results.end(); ++it)
	{
		auto P = it->get(); // get pair from future
		auto hndlFile = CreateFile(P.second.c_str(), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0); //creating handle for file
		std::wstring sizeFileStr, sizeFileStrTemp(P.second);
		BY_HANDLE_FILE_INFORMATION fileInfoStruct;
		GetFileInformationByHandle(hndlFile, &fileInfoStruct);
		
		int i = 0;
		int size = sizeFileStrTemp.size() - 1;
		SYSTEMTIME systemFileTime;
		auto strSz = fileInfoStruct.nFileIndexLow;//geting file size
			
		while (sizeFileStrTemp[size - i] != L'\\') i++; //
		sizeFileStr = sizeFileStrTemp.substr(size - i + 1);//Converting fullPath into <filename>

		FileTimeToSystemTime(&(fileInfoStruct.ftCreationTime), &systemFileTime);// coverting FILETIME to SYSTEMTIME
		//creating string to write into multiset of strings for later loging
		sizeFileStr += L"; Size: " + std::to_wstring(strSz) + L" Bytes" + L"; Date: "
			+ std::to_wstring(systemFileTime.wDay) + L"/" + std::to_wstring(systemFileTime.wMonth) + L"/"
			+ std::to_wstring(systemFileTime.wYear) + L"; Time: " + std::to_wstring(systemFileTime.wHour) + L":"
			+ std::to_wstring(systemFileTime.wMinute) + L":" + std::to_wstring(systemFileTime.wSecond)
			+ L" Per-Byte sum: " + std::to_wstring(P.first); // into string
		fileName_info.insert(sizeFileStr);
	}
		
	WriteToLogFile(fileName_info);
	
    return S_OK;
}


//
//   FUNCTION: CFileContextMenuExt::GetCommandString
//
//   PURPOSE: If a user highlights one of the items added by a context menu 
//            handler, the handler's IContextMenu::GetCommandString method is 
//            called to request a Help text string that will be displayed on 
//            the Windows Explorer status bar. This method can also be called 
//            to request the verb string that is assigned to a command. 
//            Either ANSI or Unicode verb strings can be requested. This 
//            example only implements support for the Unicode values of 
//            uFlags, because only those have been used in Windows Explorer 
//            since Windows 2000.
//
IFACEMETHODIMP FileContextMenuExt::GetCommandString(UINT_PTR idCommand, 
    UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    if (idCommand == IDM_DISPLAY)
    {
        switch (uFlags)
        {
        case GCS_HELPTEXTW:
            // Only useful for pre-Vista versions of Windows that have a 
            // Status bar.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbHelpText);
            break;

        case GCS_VERBW:
            // GCS_VERBW is an optional feature that enables a caller to 
            // discover the canonical name for the verb passed in through 
            // idCommand.
            hr = StringCchCopy(reinterpret_cast<PWSTR>(pszName), cchMax, 
                m_pwszVerbCanonicalName);
            break;

        default:
            hr = S_OK;
        }
    }

    // If the command (idCommand) is not supported by this context menu 
    // extension handler, return E_INVALIDARG.

    return hr;
}

#pragma endregion