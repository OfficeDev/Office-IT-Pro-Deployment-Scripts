<?xml version="1.0" encoding="utf-8"?>
<serviceModel xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" name="DeployOfficeProPlusSelfService" generation="1" functional="0" release="0" Id="43e7f8ec-4009-4571-a61d-02baa1df5af5" dslVersion="1.2.0.0" xmlns="http://schemas.microsoft.com/dsltools/RDSM">
  <groups>
    <group name="DeployOfficeProPlusSelfServiceGroup" generation="1" functional="0" release="0">
      <componentports>
        <inPort name="OfficeProPlusSelfServiceSite:Endpoint1" protocol="http">
          <inToChannel>
            <lBChannelMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/LB:OfficeProPlusSelfServiceSite:Endpoint1" />
          </inToChannel>
        </inPort>
        <inPort name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" protocol="tcp">
          <inToChannel>
            <lBChannelMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </inToChannel>
        </inPort>
        <inPort name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint" protocol="tcp">
          <inToChannel>
            <lBChannelMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint" />
          </inToChannel>
        </inPort>
      </componentports>
      <settings>
        <aCS name="Certificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
          </maps>
        </aCS>
        <aCS name="Certificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ClientThumbprint" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ClientThumbprint" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Enabled" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Enabled" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Version" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Version" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ServerThumbprint" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ServerThumbprint" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" defaultValue="">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" />
          </maps>
        </aCS>
        <aCS name="OfficeProPlusSelfServiceSiteInstances" defaultValue="[1,1,1]">
          <maps>
            <mapMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/MapOfficeProPlusSelfServiceSiteInstances" />
          </maps>
        </aCS>
      </settings>
      <channels>
        <sFSwitchChannel name="IE:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector" />
          </toPorts>
        </sFSwitchChannel>
        <sFSwitchChannel name="IE:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.FileUpload">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.FileUpload" />
          </toPorts>
        </sFSwitchChannel>
        <sFSwitchChannel name="IE:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Forwarder">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.Forwarder" />
          </toPorts>
        </sFSwitchChannel>
        <lBChannel name="LB:OfficeProPlusSelfServiceSite:Endpoint1">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Endpoint1" />
          </toPorts>
        </lBChannel>
        <lBChannel name="LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </toPorts>
        </lBChannel>
        <lBChannel name="LB:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint" />
          </toPorts>
        </lBChannel>
        <sFSwitchChannel name="SW:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp">
          <toPorts>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" />
          </toPorts>
        </sFSwitchChannel>
      </channels>
      <maps>
        <map name="MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" kind="Identity">
          <certificate>
            <certificateMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
          </certificate>
        </map>
        <map name="MapCertificate|OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" kind="Identity">
          <certificate>
            <certificateMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" />
          </certificate>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ClientThumbprint" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.ClientThumbprint" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Enabled" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Enabled" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Version" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Version" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteDebugger.ServerThumbprint" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.ServerThumbprint" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" kind="Identity">
          <setting>
            <aCSMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" />
          </setting>
        </map>
        <map name="MapOfficeProPlusSelfServiceSiteInstances" kind="Identity">
          <setting>
            <sCSPolicyIDMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSiteInstances" />
          </setting>
        </map>
      </maps>
      <components>
        <groupHascomponents>
          <role name="OfficeProPlusSelfServiceSite" generation="1" functional="0" release="0" software="D:\Office-IT-Pro-Deployment-Scripts\Office-ProPlus-Deployment\SelfServiceWebDeployment\Deploy\csx\Release\roles\OfficeProPlusSelfServiceSite" entryPoint="base\x64\WaHostBootstrapper.exe" parameters="base\x64\WaIISHost.exe " memIndex="-1" hostingEnvironment="frontendadmin" hostingEnvironmentVersion="2">
            <componentports>
              <inPort name="Endpoint1" protocol="http" portRanges="80" />
              <inPort name="Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" protocol="tcp" />
              <inPort name="Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint" protocol="tcp" portRanges="8172" />
              <inPort name="Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" protocol="tcp" portRanges="3389" />
              <outPort name="OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" protocol="tcp">
                <outToChannel>
                  <sFSwitchChannelMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/SW:OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp" />
                </outToChannel>
              </outPort>
            </componentports>
            <settings>
              <aCS name="Microsoft.WindowsAzure.Plugins.Diagnostics.ConnectionString" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountEncryptedPassword" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountExpiration" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.AccountUsername" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteAccess.Enabled" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteDebugger.ClientThumbprint" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Enabled" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector.Version" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteDebugger.ServerThumbprint" defaultValue="" />
              <aCS name="Microsoft.WindowsAzure.Plugins.RemoteForwarder.Enabled" defaultValue="" />
              <aCS name="__ModelData" defaultValue="&lt;m role=&quot;OfficeProPlusSelfServiceSite&quot; xmlns=&quot;urn:azure:m:v1&quot;&gt;&lt;r name=&quot;OfficeProPlusSelfServiceSite&quot;&gt;&lt;e name=&quot;Endpoint1&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteAccess.Rdp&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteDebugger.Connector&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteDebugger.FileUpload&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteDebugger.Forwarder&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput&quot; /&gt;&lt;e name=&quot;Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint&quot; /&gt;&lt;/r&gt;&lt;/m&gt;" />
            </settings>
            <resourcereferences>
              <resourceReference name="DiagnosticStore" defaultAmount="[4096,4096,4096]" defaultSticky="true" kind="Directory" />
              <resourceReference name="EventStore" defaultAmount="[1000,1000,1000]" defaultSticky="false" kind="LogStore" />
            </resourcereferences>
            <storedcertificates>
              <storedCertificate name="Stored0Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" certificateStore="My" certificateLocation="System">
                <certificate>
                  <certificateMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
                </certificate>
              </storedCertificate>
              <storedCertificate name="Stored1Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" certificateStore="My" certificateLocation="System">
                <certificate>
                  <certificateMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite/Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" />
                </certificate>
              </storedCertificate>
            </storedcertificates>
            <certificates>
              <certificate name="Microsoft.WindowsAzure.Plugins.RemoteAccess.PasswordEncryption" />
              <certificate name="Microsoft.WindowsAzure.Plugins.RemoteDebugger.TransportValidation" />
            </certificates>
          </role>
          <sCSPolicy>
            <sCSPolicyIDMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSiteInstances" />
            <sCSPolicyUpdateDomainMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSiteUpgradeDomains" />
            <sCSPolicyFaultDomainMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSiteFaultDomains" />
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
    <implementation Id="022729e5-8972-4830-b8a3-bce80e503f4c" ref="Microsoft.RedDog.Contract\ServiceContract\DeployOfficeProPlusSelfServiceContract@ServiceDefinition">
      <interfacereferences>
        <interfaceReference Id="c101ba34-2d70-4406-9052-969b39e81593" ref="Microsoft.RedDog.Contract\Interface\OfficeProPlusSelfServiceSite:Endpoint1@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite:Endpoint1" />
          </inPort>
        </interfaceReference>
        <interfaceReference Id="3e73a3be-d788-4da4-b497-2580072ebfc8" ref="Microsoft.RedDog.Contract\Interface\OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.RemoteForwarder.RdpInput" />
          </inPort>
        </interfaceReference>
        <interfaceReference Id="4c94e491-eec8-4fee-91e8-7232fc0789c0" ref="Microsoft.RedDog.Contract\Interface\OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint@ServiceDefinition">
          <inPort>
            <inPortMoniker name="/DeployOfficeProPlusSelfService/DeployOfficeProPlusSelfServiceGroup/OfficeProPlusSelfServiceSite:Microsoft.WindowsAzure.Plugins.WebDeploy.InputEndpoint" />
          </inPort>
        </interfaceReference>
      </interfacereferences>
    </implementation>
  </implements>
</serviceModel>