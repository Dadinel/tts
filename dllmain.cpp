// dllmain.cpp : Define o ponto de entrada para o aplicativo DLL.
#include "pch.h"
#include "sapi.h"
#include "string"

//Variables
ISpVoice* pVoice;
HRESULT hr;

//Functions
extern "C" __declspec(dllexport) void ExecInClientDLL(int idCommand, char* buffParam, char* buffOutput, int buffLen);
LPCWSTR charPointToLpcwStr(char* protheusText);

//Function exported to dll, read via Protheus Smartclient
extern "C" __declspec(dllexport) void ExecInClientDLL(int idCommand, char* buffParam, char* buffOutput, int buffLen)
{

	//Example
	/*
	if (FAILED(::CoInitialize(NULL)))
		return;

	HRESULT hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL,
		IID_ISpVoice, (void**)&pVoice);

	if (!SUCCEEDED(hr))
		return;

	hr = pVoice->Speak(L"Hello world", 0, NULL);
	pVoice->Release();
	pVoice = NULL;
	*/

	switch (idCommand) {
		//Create the voice object
		case 1: {
			hr = CoCreateInstance(CLSID_SpVoice, NULL, CLSCTX_ALL,
				IID_ISpVoice, (void**)&pVoice);
			break;
		}

		//Speech with SPF_PURGEBEFORESPEAK
		case 2: {
			pVoice->Speak(charPointToLpcwStr(buffParam), SPF_PURGEBEFORESPEAK, NULL);
			break;
		}

		//Clears the voice object
		case 3: {
			pVoice->Release();
			pVoice = NULL;
			hr = NULL;
			::CoUninitialize();
			break;
		}

		//Sets rate
		case 4: {
			pVoice->SetRate(atol(buffParam));
			break;
		}

		//Change the volume
		case 5: {
			pVoice->SetVolume(atof(buffParam));
			break;
		}

		//Speech with SPF_ASYNC
		case 6: {
			std::string str(buffParam);

			std::wstring stemp = std::wstring(str.begin(), str.end());
			LPCWSTR sw = stemp.c_str();

			pVoice->Speak(charPointToLpcwStr(buffParam), SPF_ASYNC, NULL);
			break;
		}
	}
}

//Covert char* to LPCWSTR 
LPCWSTR charPointToLpcwStr(char* protheusText) {
	std::string stringProtheusText(protheusText);

	std::wstring stemp = std::wstring(stringProtheusText.begin(), stringProtheusText.end());
	LPCWSTR lpcwStrProthuesText = stemp.c_str();

	return lpcwStrProthuesText;
}