// CppComClient.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include "pch.h"
#include <iostream>
#include <Windows.h>

#import "C:\build\ComInterop\DotNetComLibrary\bin\Debug\DotNetComLibrary.tlb" named_guids raw_interfaces_only


class EventSink : public DotNetComLibrary::IEventSink
{
private:
	long	m_lReferenceCount;

public:
	EventSink()
	{
		m_lReferenceCount = 0;
	}

	virtual HRESULT __stdcall OnEvent(/*[in]*/ long i)
	{
		std::cout << "got event: " << i << "\n";
		return S_OK;
	}

	std::string getIIDName(REFIID riid)
	{
		if (riid == IID_IDispatch) return "IDispatch";
		else if (riid == IID_IErrorInfo) return "IErrorInfo";
		else if (riid == IID_IProvideClassInfo) return "IID_IProvideClassInfo";
		else if (riid == IID_IProvideClassInfo2) return "IID_IProvideClassInfo2";
		else if (riid == IID_ISupportErrorInfo) return "IID_ISupportErrorInfo";
		else if (riid == IID_ITypeInfo) return "IID_ITypeInfo";
		else if (riid == IID_IUnknown) return "IID_IUnknown";
		else if (riid == IID_IConnectionPoint) return "IID_IConnectionPoint";
		else if (riid == IID_IEnumVARIANT) return "IID_IEnumVARIANT";
		else if (riid == IID_IMoniker) return "IID_IMoniker";

		return "unknown?";
	}

	// Inherited via IEventSink
	virtual HRESULT __stdcall QueryInterface(REFIID riid, void ** ppvObject) override
	{
		LPOLESTR lplpsz;
		StringFromIID(riid, &lplpsz);


		// from here: https://docs.microsoft.com/en-us/dotnet/framework/interop/com-callable-wrapper
		// are we being called with one of these interfaces?
		/*
Interface	Description
Idispatch	Provides a mechanism for late binding to type.
IerrorInfo	Provides a textual description of the error, its source, a Help file, Help context, and the GUID of the interface that defined the error (always GUID_NULL for .NET classes).
IprovideClassInfo	Enables COM clients to gain access to the ITypeInfo interface implemented by a managed class.
IsupportErrorInfo	Enables a COM client to determine whether the managed object supports the IErrorInfo interface. If so, enables the client to obtain a pointer to the latest exception object. All managed types support the IErrorInfo interface.
ItypeInfo	Provides type information for a class that is exactly the same as the type information produced by Tlbexp.exe.
Iunknown	Provides the standard implementation of the IUnknown interface with which the COM client manages the lifetime of the CCW and provides type coercion.
A managed class can also provide the COM interfaces described in the following table.

Interface	Description
The (_classname) class interface	Interface, exposed by the runtime and not explicitly defined, that exposes all public interfaces, methods, properties, and fields that are explicitly exposed on a managed object.
IConnectionPoint and IconnectionPointContainer	Interface for objects that source delegate-based events (an interface for registering event subscribers).
IdispatchEx	Interface supplied by the runtime if the class implements IExpando. The IDispatchEx interface is an extension of the IDispatch interface that, unlike IDispatch, enables enumeration, addition, deletion, and case-sensitive calling of members.
IEnumVARIANT	Interface for collection-type classes, which enumerates the objects in the collection if the class implements IEnumerable.
		*/

		std::wstring interfaceName = lplpsz;
		std::cout << "QueryInterface (iid=" << interfaceName.c_str() << ")\n";
		std::cout << "requesting iid of: " << getIIDName(riid).c_str() << "\n";

		

		if (riid == DotNetComLibrary::IID_IEventSink)
		{
			std::cout << "QueryInterface is IEventSink\n";
			*ppvObject = this;
			AddRef();
			return S_OK;
		}
		else if (riid == IID_IUnknown)
		{
			std::cout << "QueryInterface is IUnknown\n";
			*ppvObject = this;
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}
	virtual ULONG __stdcall AddRef(void) override
	{
		std::cout << "AddRef\n";
		return InterlockedIncrement(&m_lReferenceCount);
	}
	virtual ULONG __stdcall Release(void) override
	{
		std::cout << "Release\n";
		if (InterlockedDecrement(&m_lReferenceCount) == 0)
		{
			std::cout << "deleting this\n";
			delete this;
			return 0;
		}
		return m_lReferenceCount;
	}

	virtual HRESULT __stdcall GetTypeInfoCount(UINT * pctinfo) override
	{
		std::cout << "GetTypeInfoCount\n";
		return E_FAIL;
	}
	virtual HRESULT __stdcall GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo) override
	{
		std::cout << "GetTypeInfo\n";
		return E_FAIL;
	}
	virtual HRESULT __stdcall GetIDsOfNames(REFIID riid, LPOLESTR * rgszNames, UINT cNames, LCID lcid, DISPID * rgDispId) override
	{
		std::cout << "GetIDsOfNames\n";
		return E_FAIL;
	}
	virtual HRESULT __stdcall Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, UINT * puArgErr) override
	{
		std::cout << "Invoke\n";
		return E_FAIL;
	}
};

int main()
{
    std::cout << "Hello World!\n"; 

	CoInitialize(NULL);

	auto pEventSink = new EventSink();

	DotNetComLibrary::IServerObjectPtr pServerObject;


	HRESULT hr = pServerObject.CreateInstance(DotNetComLibrary::CLSID_ServerObject);
	if (hr == S_OK)
	{
		std::cout << "Call RegisterEventSink\n";
		hr = pServerObject->RegisterEventSink(pEventSink);
		std::cout << "hr=" << hr << "\n";

		/*
			Calling RegisterEventSink results in lots of calls being made to QueryInterface with different IIDs
			which is confusing?
			unless it's because my AddRef/Release aren't implemented yet.

			This is sample output:
c:\build\ComInterop\Debug>CppComClient.exe
Hello World!
Call RegisterEventSink
QueryInterface (iid=00133578)
requesting iid of: IID_IUnknown
QueryInterface is IUnknown
QueryInterface (iid=00133520)
requesting iid of: unknown?
QueryInterface (iid=001335D0)
requesting iid of: IID_IProvideClassInfo
QueryInterface (iid=001334C8)
requesting iid of: unknown?
AddRef
QueryInterface (iid=00132FF8)
requesting iid of: unknown?
QueryInterface (iid=00133310)
requesting iid of: unknown?
QueryInterface (iid=00133628)
requesting iid of: unknown?
QueryInterface (iid=001330A8)
requesting iid of: unknown?
Release
QueryInterface (iid=00133158)
requesting iid of: unknown?
QueryInterface is IEventSink
AddRef
Release
hr=0
Call AddNumber
hr=0 result=93
Call ShowWindow
QueryInterface (iid=00133680)
requesting iid of: unknown?
QueryInterface (iid=00133418)
requesting iid of: unknown?
QueryInterface (iid=001336D8)
requesting iid of: unknown?
QueryInterface (iid=00133730)
requesting iid of: unknown?
QueryInterface (iid=00133788)
requesting iid of: IID_IUnknown
QueryInterface is IUnknown
AddRef
QueryInterface (iid=00132D90)
requesting iid of: unknown?
QueryInterface (iid=00133208)
requesting iid of: unknown?
QueryInterface (iid=001331B0)
requesting iid of: unknown?
QueryInterface (iid=00132E40)
requesting iid of: unknown?
QueryInterface (iid=00132DE8)
requesting iid of: unknown?
QueryInterface (iid=00133050)
requesting iid of: unknown?
QueryInterface (iid=00132E98)
requesting iid of: unknown?
QueryInterface (iid=00132EF0)
requesting iid of: unknown?
QueryInterface (iid=00132F48)
requesting iid of: unknown?
QueryInterface (iid=00132FA0)
requesting iid of: unknown?
QueryInterface (iid=001333C0)
requesting iid of: unknown?
QueryInterface (iid=00133940)
requesting iid of: unknown?
QueryInterface (iid=00133838)
requesting iid of: unknown?
QueryInterface (iid=001338E8)
requesting iid of: unknown?
QueryInterface (iid=00133C00)
requesting iid of: unknown?
QueryInterface (iid=00133A48)
requesting iid of: unknown?
QueryInterface (iid=00133BA8)
requesting iid of: unknown?
QueryInterface (iid=00133AA0)
requesting iid of: unknown?
QueryInterface (iid=00133B50)
requesting iid of: unknown?
QueryInterface (iid=00133AF8)
requesting iid of: unknown?
QueryInterface (iid=00133998)
requesting iid of: unknown?
QueryInterface (iid=00133C58)
requesting iid of: unknown?
QueryInterface (iid=001339F0)
requesting iid of: unknown?
QueryInterface (iid=00133CB0)
requesting iid of: unknown?
QueryInterface (iid=00133890)
requesting iid of: unknown?
Release
QueryInterface (iid=0016ACD8)
requesting iid of: unknown?
QueryInterface is IEventSink
AddRef
QueryInterface (iid=0016AC80)
requesting iid of: unknown?
QueryInterface (iid=0016AAC8)
requesting iid of: unknown?
QueryInterface is IEventSink
got event: 0
got event: 1
hr=0
Call ThrowException
hr=-2146233088
		
		*/



		std::cout << "Call AddNumber\n";

		long result = 0;
		hr = pServerObject->AddARandomNumber(10, &result);
		std::cout << "hr=" << hr << " result=" << result << "\n";

		std::cout << "Call ShowWindow\n";
		BSTR message = ::SysAllocString(L"testing, testing, 1, 2, 3");
		hr = pServerObject->ShowMessage(message);
		::SysFreeString(message);
		std::cout << "hr=" << hr << "\n";

		std::cout << "Call ThrowException\n";
		hr = pServerObject->ThrowAnException();
		std::cout << "hr=" << hr << "\n";

	}
	else
	{
		std::cout << "Failed to create com object\n";
	}

	CoUninitialize();
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
