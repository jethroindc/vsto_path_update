## Wix Custom Action to Update VSTO Location Dynamically

Recently we had to deploy a MS Word template along with a VSTO component.  The only problem is that the MS Word template needed to point to the VSTO component at runtime and there is no way to handle this dynamically.

Thus we created this custom action to help set the VSTO location at install time ensuring everyone plays together nicely.

Quick snippet of how this could be used in a WIX xml definition 
```xml
  <Product ...>

    ...

    <InstallExecuteSequence>
      <Custom Action='CustomValue' Before='UpdateWordTemplate'/>
      <Custom Action='UpdateWordTemplate' Before='InstallFinalize'>NOT Installed AND NOT REMOVE</Custom>
    </InstallExecuteSequence>
  </Product>
  <Fragment>
    <CustomAction Id="CustomValue" Property="UpdateWordTemplate" Value="TEMPLATE=[#memo_template];VSTO=[#Memo_vsto]"/>
    <CustomAction Id='UpdateWordTemplate' BinaryKey='UpdateWordTemplate' DllEntry='UpdateAddonPath' Impersonate='no' Execute='deferred' Return='check' />
    <Binary Id='UpdateWordTemplate' SourceFile='..\UpdateWordTemplateCustomAction\bin\Debug\UpdateWordTemplateCustomAction.CA.dll'/>
  </Fragment>
```
