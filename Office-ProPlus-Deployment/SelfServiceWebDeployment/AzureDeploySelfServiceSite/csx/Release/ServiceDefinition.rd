<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="AzureDeploySelfServiceSite" generation="1" functional="0" release="0" Id="14672e9d-d3f2-461b-a614-774366b2fb85" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
  <groups>
    <group name="AzureDeploySelfServiceSiteGroup" generation="1" functional="0" release="0">
      <componentports>
        <inPort name="OfficeProPlusSelfServiceSite:Endpoint1" protocol="http">
          <inToChannel>
            <lBChannelMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/LB:OfficeProPlusSelfServiceSite:Endpoint1" />
          </inToChannel>
        </inPort>
        <inPort name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" protocol="tcp">
          <inToChannel>
            <lBChannelMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </inToChannel>
        </inPort>
      </componentports>
      <settings>
        <aCS name="Certificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" defaultValue="">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSiteInstances" defaultValue="[1,1,1]">
          <maps>
            <mapMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/MapOfficeProPlusSelfServiceSiteInstances" />
          </maps>
        </aCS>
      </settings>
      <channels>
        <lBChannel name="LB:OfficeProPlusSelfServiceSite:Endpoint1">
          <toPorts>
            <inPortMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Endpoint1" />
          </toPorts>
        </lBChannel>
        <lBChannel name="LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput">
          <toPorts>
            <inPortMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </toPorts>
        </lBChannel>
        <sFSwitchChannel name="SW:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp">
          <toPorts>
            <inPortMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" />
          </toPorts>
        </sFSwitchChannel>
      </channels>
      <maps>
        <map name="MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" kind="Identity">
          <certificate>
            <certificateMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
          </certificate>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" kind="Identity">
          <setting>
            <aCSMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSiteInstances" kind="Identity">
          <setting>
            <sCSPolicyIDMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSiteInstances" />
          </setting>
        </map>
      </maps>
      <components>
        <groupHascomponents>
          <role name="OfficeProPlusSelfServiceSite" generation="1" functional="0" release="0" software="D:\Office-IT-Pro-Deployment-Scripts\Office-ProPlus-Deployment\SelfServiceWebDeployment\AzureDeploySelfServiceSite\csx\Release\roles\OfficeProPlusSelfServiceSite" entryPoint="base\x64\WaHostBootstrapper.exe" parameters="base\x64\WaIISHost.exe " memIndex="-1" hostingEnvironment="frontendadmin" hostingEnvironmentVersion="2">
            <componentports>
              <inPort name="Endpoint1" protocol="http" portRanges="80" />
              <inPort name="Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" protocol="tcp" />
              <inPort name="Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" protocol="tcp" portRanges="3389" />
              <outPort name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" protocol="tcp">
                <outToChannel>
                  <sFSwitchChannelMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/SW:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" />
                </outToChannel>
              </outPort>
            </componentports>
            <settings>
              <aCS name="Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" defaultValue="" />
              <aCS name="__ModelData" defaultValue="&lt;m role=&quot;OfficeProPlusSelfServiceSite&quot; xmlns=&quot;urn:azure:m:v1&quot;&gt;&lt;r name=&quot;OfficeProPlusSelfServiceSite&quot;&gt;&lt;e name=&quot;Endpoint1&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput&quot; /&gt;&lt;/r&gt;&lt;/m&gt;" />
            </settings>
            <resourcereferences>
              <resourceReference name="DiagnosticStore" defaultAmount="[4096,4096,4096]" defaultSticky="true" kind="Directory" />
              <resourceReference name="EventStore" defaultAmount="[1000,1000,1000]" defaultSticky="false" kind="LogStore" />
            </resourcereferences>
            <storedcertificates>
              <storedCertificate name="Stored0Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" certificateStore="My" certificateLocation="System">
                <certificate>
                  <certificateMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
                </certificate>
              </storedCertificate>
            </storedcertificates>
            <certificates>
              <certificate name="Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
            </certificates>
          </role>
          <sCSPolicy>
            <sCSPolicyIDMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSiteInstances" />
            <sCSPolicyUpdateDomainMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSiteUpgradeDomains" />
            <sCSPolicyFaultDomainMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSiteFaultDomains" />
          </sCSPolicy>
        </groupHascomponents>
      </components>
      <sCSPolicy>
        <sCSPolicyUpdateDomain name="OfficeProPlusSelfServiceSiteUpgradeDomains" defaultPolicy="[5,5,5]" />
        <sCSPolicyFaultDomain name="OfficeProPlusSelfServiceSiteFaultDomains" defaultPolicy="[2,2,2]" />
        <sCSPolicyID name="OfficeProPlusSelfServiceSiteInstances" defaultPolicy="[1,1,1]" />
      </sCSPolicy>
    </group>
  </groups>
  <implements>
    <implementation Id="95eda96b-fb60-4f93-b781-89a0b36c25fd" ref="Microsoft.RedDog.Contract\ServiceContract\AzureDeploySelfServiceSiteContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="e736e4b1-3a06-4b18-973a-4de436d9d1ed" ref="Microsoft.RedDog.Contract\Interface\OfficeProPlusSelfServiceSite:Endpoint1@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite:Endpoint1" />
          </inPort>
        </interfaceReference>
        <interfaceReference Id="a77431a0-66ad-46e8-ad9c-bd9b31ce66ec" ref="Microsoft.RedDog.Contract\Interface\OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/AzureDeploySelfServiceSite/AzureDeploySelfServiceSiteGroup/OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>