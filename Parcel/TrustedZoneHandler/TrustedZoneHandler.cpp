#define STRICT
#include "stdafx.h"

extern "C" {
	// more or less borrowed from Raymond Chen's page on
	// reading, setting, and deleting security identifiers:
	// http://blogs.msdn.com/b/oldnewthing/archive/2013/11/04/10463035.aspx

	extern "C" __declspec(dllexport) void EraseZoneData(wchar_t *);
	extern "C" __declspec(dllexport) DWORD ReadZoneData(wchar_t *);
	extern "C" __declspec(dllexport) void SetZoneData(wchar_t *, DWORD);

	void EraseZoneData(wchar_t *filename)
	{
		// start up COM
		HRESULT hr = CoInitialize(NULL);

		// COM smart ptr to zone ID
		CComPtr<IZoneIdentifier> spzi;
		// create zone ID object
		spzi.CoCreateInstance(CLSID_PersistentZoneIdentifier);
		// remove zone information
		spzi->Remove();
		/// overwrite zone information on file
		CComQIPtr<IPersistFile>(spzi)->Save(filename, TRUE);

		// shutdown COM
		CoUninitialize();
	}

	DWORD ReadZoneData(wchar_t *filename)
	{
		// start up COM
		HRESULT hr = CoInitialize(NULL);

		// COM smart pointer to security mgr interface
		CComPtr<IInternetSecurityManager> spzi;
		spzi.CoCreateInstance(CLSID_InternetSecurityManager);

		// read security DWORD; return URLZONE_INVALID if
		// this call fails
		DWORD dwZone;
		if (SUCCEEDED(spzi->MapUrlToZone(
						    filename,
						    &dwZone,
						    MUTZ_ISFILE | MUTZ_DONT_UNESCAPE))) {
			return dwZone;
		} else {
			return URLZONE_INVALID;
		}

		// shutdown COM
		CoUninitialize();
	}

	void SetZoneData(wchar_t *filename, DWORD urlz)
	{
		// start up COM
		HRESULT hr = CoInitialize(NULL);

		// COM smart ptr to zone ID
		CComPtr<IZoneIdentifier> spzi;
		// create zone ID object
		spzi.CoCreateInstance(CLSID_PersistentZoneIdentifier);
		// set field to user-defined url zone
		spzi->SetId(urlz);
		// overwrite zone information on file
		CComQIPtr<IPersistFile>(spzi)->Save(filename, TRUE);
		
		// shutdown COM
		CoUninitialize();
	}
}