#include <windows.h>

#include <propsys.h>     // Property System APIs and interfaces
#include <propkey.h>     // System PROPERTYKEY definitions
#include <propvarutil.h> // PROPVARIANT and VARIANT helper APIs

class CRecipePropertyHandler :
    public IPropertyStore,// �˽ӿڹ�������ö�ٺͲ�������ֵ�ķ�����
    public IPropertyStoreCapabilities,// ����һ���������÷���ȷ���û��Ƿ������UI�б༭���ԡ�
    public IInitializeWithStream// ʹ����������ʼ���������(�����Դ����������ͼ��������Ԥ���������)�ķ�����
{
public:
    CRecipePropertyHandler() : _cRef(0) { }

    // IUnknown 
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
    {
        //���������ÿһ����Ҫ��¶��ȥ�ĵ�com�ӿ�,���
        static const QITAB qit[] = {
            QITABENT(CRecipePropertyHandler, IPropertyStore),
            QITABENT(CRecipePropertyHandler, IPropertyStoreCapabilities),
            QITABENT(CRecipePropertyHandler, IInitializeWithStream),
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&_cRef); }
    IFACEMETHODIMP_(ULONG) Release()
    {
        ULONG cRef = InterlockedDecrement(&_cRef);
        if (!cRef) { delete this; }
        return cRef;
    }

    // IPropertyStore
    // �˽ӿڹ�������ö�ٺͲ�������ֵ�ķ�����

    //�˷������ظ��ӵ��ļ����������ļ�����
    IFACEMETHODIMP GetCount(DWORD* pcProps) { *pcProps = 0; return _pCache_IPropertyStoreCache_���� ? _pCache_IPropertyStoreCache_����->GetCount(pcProps) : E_UNEXPECTED; }
    //��������������л�ȡ���Լ���
    IFACEMETHODIMP GetAt(DWORD iProp, PROPERTYKEY* pkey) { *pkey = PKEY_Null; return _pCache_IPropertyStoreCache_���� ? _pCache_IPropertyStoreCache_����->GetAt(iProp, pkey) : E_UNEXPECTED; }
    IFACEMETHODIMP GetValue(REFPROPERTYKEY key, PROPVARIANT* pPropVar);//�˷��������ض����Ե����ݡ�
    IFACEMETHODIMP SetValue(REFPROPERTYKEY key, REFPROPVARIANT propVar);//�˷�����������ֵ���滻��ɾ������ֵ��
    IFACEMETHODIMP Commit();//�ڽ����˸���֮�󣬴˷���������ġ�

    // IPropertyStoreCapabilities
    // ����һ���������÷���ȷ���û��Ƿ������UI�б༭���ԡ�
    IFACEMETHODIMP IsPropertyWritable(REFPROPERTYKEY key);

    // IInitializeWithStream
    // ʹ����������ʼ���������(�����Դ����������ͼ��������Ԥ���������)�ķ�����
    IFACEMETHODIMP Initialize(IStream* pStream, DWORD grfMode);

private:
    long _cRef;
    IPropertyStoreCache* _pCache_IPropertyStoreCache_����;  // �ڲ�ֵ�����DOM��˳���IPropertyStore����
    IStream* _pStream_����Initialize��������; // ����Initialize���������������ύʱ����

    // ����ʱִ��
    ~CRecipePropertyHandler(){}
};