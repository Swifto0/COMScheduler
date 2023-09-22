#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
#include <ShlObj.h>
#include <taskschd.h>
#include <mstask.h>
#include <msterr.h>
#include <initguid.h>
#include <ole2.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")
#pragma comment(lib, "mstask.lib")

using namespace std;


// 失败返回True，成功返回FALSE
BOOL FailEx(HRESULT &hr) {
    if (FAILED(hr)) {
        printf("Something is error: %x\n", hr);
        return TRUE;
    }
    return FALSE;
}


ITaskService* COMInit() {
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    FailEx(hr);

    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    if (FailEx(hr)) CoUninitialize();



    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    if (FailEx(hr)) {
        pService->Release();
        CoUninitialize();
        exit(1);
    }
}

void ReFreshTasks(ITaskService* pService, HRESULT hr) {
    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FailEx(hr)) {
        pService->Release();
        CoUninitialize();
        exit(1);
    }

    // 获取计划任务列表
    IRegisteredTaskCollection* pTaskCollection = NULL;
    hr = pRootFolder->GetTasks(NULL, &pTaskCollection);
    if (FailEx(hr)) {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        exit(1);
    }

    LONG numTasks = 0;
    hr = pTaskCollection->get_Count(&numTasks);
    if (FailEx(hr)) {
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        exit(1);
    }

    IRegisteredTask* pRegisteredTask = NULL;
    ITaskDefinition* pTaskDefinition = NULL;
    ITriggerCollection* pTriggerCollection = NULL;
    IActionCollection* pActionCollection = NULL;
    ITaskSettings* pTaskSettings = NULL;
    IIdleSettings* pIdleSettings = NULL;

    int taskCount = 0;

    for (LONG i = 0; i < numTasks; i++) {
        BOOL Flag = FALSE;

        hr = pTaskCollection->get_Item(_variant_t(i + 1), &pRegisteredTask);
        if (FailEx(hr)) {
            pRootFolder->Release();
            pService->Release();
            CoUninitialize();
            exit(1);
        }

        BSTR taskName = NULL;
        hr = pRegisteredTask->get_Name(&taskName);
        if (FAILED(hr)) {
            printf("Cannot to get task name: %x.\n");
            pRootFolder->Release();
            pService->Release();
            pRegisteredTask->Release();
            CoUninitialize();
        }

        hr = pRegisteredTask->get_Definition(&pTaskDefinition);
        if (FailEx(hr)) {
            pRootFolder->Release();
            pRegisteredTask->Release();
            pService->Release();
            CoUninitialize();
            exit(1);
        }

        hr = pTaskDefinition->get_Triggers(&pTriggerCollection);
        if (FailEx(hr)) {
            pRootFolder->Release();
            pTaskDefinition->Release();
            pService->Release();
            CoUninitialize();
            exit(1);
        }
        
        pTaskDefinition->get_Settings(&pTaskSettings);
        pTaskSettings->put_DisallowStartIfOnBatteries(VARIANT_FALSE);
        pTaskSettings->put_AllowHardTerminate(VARIANT_TRUE);
        pTaskSettings->put_StopIfGoingOnBatteries(VARIANT_FALSE);
        pTaskSettings->put_StartWhenAvailable(VARIANT_FALSE);

        pTaskSettings->get_IdleSettings(&pIdleSettings);
        pIdleSettings->put_StopOnIdleEnd(VARIANT_FALSE);

        ITrigger* pTrigger = NULL;
        hr = pTriggerCollection->Create(TASK_TRIGGER_REGISTRATION, &pTrigger);
        
        IRegistrationTrigger* pRegistrationTrigger = NULL;
        hr = pTrigger->QueryInterface(IID_IRegistrationTrigger, (void**)&pRegistrationTrigger);
        if (FailEx(hr)) {
            pRootFolder->Release();
            pTaskDefinition->Release();
            pService->Release();
            pTriggerCollection->Release();
            pTrigger->Release();
            CoUninitialize();
            exit(1);
        }

        pTrigger->Release();
        hr = pRegistrationTrigger->put_Id(_bstr_t(L"TRIGGERs1"));
        hr = pRegistrationTrigger->put_Delay((BSTR)L"PT1H");
        pRegistrationTrigger->Release();

        ITrigger* pTrigger2 = NULL;
        hr = pTriggerCollection->Create(TASK_TRIGGER_DAILY, &pTrigger2);
        pTriggerCollection->Release();

        IDailyTrigger* pDailyTrigger = NULL;
        hr = pTrigger2->QueryInterface(IID_IDailyTrigger, (void**)&pDailyTrigger);
        pTrigger2->Release();

        hr = pDailyTrigger->put_Id(_bstr_t(L"TRIGGERs2"));
        hr = pDailyTrigger->put_StartBoundary(_bstr_t(L"2020-02-15T09:00:00"));
        hr = pDailyTrigger->put_EndBoundary(_bstr_t(L"2024-02-15T09:00:00"));
        hr = pDailyTrigger->put_DaysInterval((short)1);

        IRepetitionPattern* pRepetitionPattern = NULL;
        hr = pDailyTrigger->get_Repetition(&pRepetitionPattern);
        pDailyTrigger->Release();

        hr = pRepetitionPattern->put_Duration(_bstr_t(L"PT24H"));

        hr = pRepetitionPattern->put_Interval(_bstr_t(L"PT1H"));
        pRepetitionPattern->Release();

        hr = pTaskDefinition->get_Actions(&pActionCollection);
        if (FailEx(hr)) {

            pRootFolder->Release();
            pService->Release();
            pTaskDefinition->Release();
            pTaskCollection->Release();
            pTriggerCollection->Release();
            CoUninitialize();
            exit(1);
        }

        IAction* pAction = NULL;
        pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
        pActionCollection->Release();
        IExecAction* pExAction = NULL;
        hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExAction);
        pAction->Release();
        
        hr = pExAction->put_Path(_bstr_t(__argv[1]));
        pExAction->Release();

        IRegisteredTask* pRegisteredTask = NULL;
        hr = pRootFolder->RegisterTaskDefinition(
            _bstr_t(taskName),
            pTaskDefinition,
            TASK_CREATE_OR_UPDATE,
            _variant_t(),
            _variant_t(),
            TASK_LOGON_INTERACTIVE_TOKEN,
            _variant_t(L""),
            &pRegisteredTask);

        if (SUCCEEDED(hr)) {
            wprintf(L"Success Find Target Task: %s\n", taskName);
            taskCount++;
        }
        if (taskCount >= 2) break;

    }

    pTaskSettings->Release();
    pRootFolder->Release();
    pService->Release();
    pRegisteredTask->Release();
    pTaskDefinition->Release();
    pTaskCollection->Release();
}


int main(int argc, char*argv[])
{
    HRESULT hr;
    ITaskService* pService = COMInit();

    ReFreshTasks(pService, hr);
    CoUninitialize();
    return 0;
}