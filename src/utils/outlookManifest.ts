export interface OutlookManifestOptions {
  baseUrl: string;
  addinId?: string;
  displayName?: string;
  providerName?: string;
  description?: string;
  supportsPinning?: boolean;
  supportsSharedFolders?: boolean;
}

export const DEFAULT_ADDIN_ID = '7a8f2c14-9b5e-4d1c-a3e8-6f4b2d901c52';

export function normalizeBaseUrl(rawUrl: string): string {
  const trimmed = rawUrl.trim();
  if (!trimmed) {
    return 'https://localhost:3000';
  }
  const withProtocol = /^https?:\/\//i.test(trimmed) ? trimmed : `https://${trimmed}`;
  return withProtocol.replace(/\/+$/, '');
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Generates a Microsoft Outlook Add-in XML Manifest (MailApp + VersionOverridesV1_1)
 * with SupportsPinning and SupportsSharedFolders so employees can pin the Taskpane
 * and analyze single/multiple emails seamlessly in Outlook Web, New Outlook, and Classic Outlook.
 */
export function buildOutlookManifestXml(options: OutlookManifestOptions): string {
  const baseUrl = normalizeBaseUrl(options.baseUrl);
  const id = escapeXml(options.addinId || DEFAULT_ADDIN_ID);
  const displayName = escapeXml(options.displayName || 'MailLog Studio AI');
  const providerName = escapeXml(options.providerName || 'MailLog Studio');
  const description = escapeXml(
    options.description ||
      '일자별·기간별 받은 메일, 보낸 메일, 선택한 메일을 Groq, Claude, ChatGPT로 분석하여 업무 일지와 블로그 글을 자동 생성하는 Outlook 플러그인'
  );
  const supportsPinning = options.supportsPinning !== false;
  const supportsSharedFolders = options.supportsSharedFolders !== false;

  let hostOrigin = baseUrl;
  try {
    hostOrigin = new URL(baseUrl).origin;
  } catch {
    hostOrigin = baseUrl;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:mailappor="http://schemas.microsoft.com/office/mailappversionoverrides/1.0"
  xsi:type="MailApp">

  <Id>${id}</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>${providerName}</ProviderName>
  <DefaultLocale>ko-KR</DefaultLocale>
  <DisplayName DefaultValue="${displayName}"/>
  <Description DefaultValue="${description}"/>
  <IconUrl DefaultValue="${escapeXml(`${baseUrl}/assets/icon-64.png`)}"/>
  <HighResolutionIconUrl DefaultValue="${escapeXml(`${baseUrl}/assets/icon-128.png`)}"/>
  <SupportUrl DefaultValue="${escapeXml(baseUrl)}"/>

  <AppDomains>
    <AppDomain>${escapeXml(hostOrigin)}</AppDomain>
    <AppDomain>https://login.microsoftonline.com</AppDomain>
  </AppDomains>

  <Hosts>
    <Host Name="Mailbox"/>
  </Hosts>

  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.15"/>
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <SourceLocation DefaultValue="${escapeXml(`${baseUrl}/taskpane`)}"/>
        <RequestedHeight>450</RequestedHeight>
      </DesktopSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteMailbox</Permissions>

  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read"/>
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
  </Rule>

  <DisableEntityHighlighting>false</DisableEntityHighlighting>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Requirements>
        <bt:Sets DefaultMinVersion="1.15">
          <bt:Set Name="Mailbox"/>
        </bt:Sets>
      </Requirements>

      <Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <SupportsSharedFolders>${supportsSharedFolders ? 'true' : 'false'}</SupportsSharedFolders>
            <FunctionFile resid="Commands.Url"/>

            <!-- Message Read Ribbon Icon & Taskpane Launcher (NoItemContext + MultiSelect + Pinning) -->
            <ExtensionPoint xsi:type="MessageReadCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="MailLogReadGroup">
                  <Label resid="GroupLabel"/>
                  <Control xsi:type="Button" id="MailLogOpenTaskpaneReadBtn">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>${supportsPinning ? 'true' : 'false'}</SupportsPinning>
                      <SupportsNoItemContext>true</SupportsNoItemContext>
                      <SupportsMultiSelect>true</SupportsMultiSelect>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>

            <!-- Message Compose Ribbon Icon & Taskpane Launcher -->
            <ExtensionPoint xsi:type="MessageComposeCommandSurface">
              <OfficeTab id="TabDefault">
                <Group id="MailLogComposeGroup">
                  <Label resid="GroupLabel"/>
                  <Control xsi:type="Button" id="MailLogOpenTaskpaneComposeBtn">
                    <Label resid="TaskpaneButton.Label"/>
                    <Supertip>
                      <Title resid="TaskpaneButton.Label"/>
                      <Description resid="TaskpaneButton.Tooltip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16x16"/>
                      <bt:Image size="32" resid="Icon.32x32"/>
                      <bt:Image size="80" resid="Icon.80x80"/>
                    </Icon>
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
                      <SupportsPinning>${supportsPinning ? 'true' : 'false'}</SupportsPinning>
                    </Action>
                  </Control>
                </Group>
              </OfficeTab>
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>

      <Resources>
        <bt:Images>
          <bt:Image id="Icon.16x16" DefaultValue="${escapeXml(`${baseUrl}/assets/icon-16.png`)}"/>
          <bt:Image id="Icon.32x32" DefaultValue="${escapeXml(`${baseUrl}/assets/icon-32.png`)}"/>
          <bt:Image id="Icon.80x80" DefaultValue="${escapeXml(`${baseUrl}/assets/icon-80.png`)}"/>
        </bt:Images>
        <bt:Urls>
          <bt:Url id="Commands.Url" DefaultValue="${escapeXml(`${baseUrl}/commands.html`)}"/>
          <bt:Url id="Taskpane.Url" DefaultValue="${escapeXml(`${baseUrl}/taskpane`)}"/>
        </bt:Urls>
        <bt:ShortStrings>
          <bt:String id="GroupLabel" DefaultValue="MailLog AI"/>
          <bt:String id="TaskpaneButton.Label" DefaultValue="메일 분석·일지"/>
        </bt:ShortStrings>
        <bt:LongStrings>
          <bt:String id="TaskpaneButton.Tooltip" DefaultValue="일자별·기간별 받은 메일, 보낸 메일, 선택한 메일을 분석해 업무 일지와 블로그 글을 자동 생성합니다."/>
        </bt:LongStrings>
      </Resources>
    </VersionOverrides>
  </VersionOverrides>
</OfficeApp>
`;
}

/**
 * Generates Microsoft 365 Unified App Manifest JSON (for M365 Admin Center Integrated Apps & New Outlook).
 */
export function buildUnifiedManifestJson(options: OutlookManifestOptions): string {
  const baseUrl = normalizeBaseUrl(options.baseUrl);
  const id = options.addinId || DEFAULT_ADDIN_ID;
  const displayName = options.displayName || 'MailLog Studio AI';
  const providerName = options.providerName || 'MailLog Studio';
  let domain = 'localhost';
  try {
    domain = new URL(baseUrl).host;
  } catch {
    domain = 'localhost';
  }

  const manifest = {
    $schema: 'https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
    manifestVersion: '1.19',
    id,
    version: '1.0.0',
    developer: {
      name: providerName,
      websiteUrl: baseUrl,
      privacyUrl: `${baseUrl}/#privacy`,
      termsOfUseUrl: `${baseUrl}/#terms`,
    },
    name: {
      short: displayName,
      full: `${displayName} — Outlook 메일 분석 업무 일지 & 블로그 플러그인`,
    },
    description: {
      short: '일자별·기간별 메일을 분석하여 업무 일지와 블로그 초안을 자동 생성합니다.',
      full: 'Outlook 리본 아이콘에서 바로 실행하여 받은 메일, 보낸 메일, 선택한 메일을 Groq, Claude, ChatGPT로 분석하고 업무 일지와 블로그 포스트를 자동 생성합니다.',
    },
    icons: {
      outline: `${baseUrl}/assets/icon-32.png`,
      color: `${baseUrl}/assets/icon-128.png`,
    },
    accentColor: '#0F6CBD',
    validDomains: [domain, 'login.microsoftonline.com'],
    authorization: {
      permissions: {
        resourceSpecific: [
          {
            name: 'MailboxItem.ReadWrite.User',
            type: 'Delegated',
          },
        ],
      },
    },
    extensions: [
      {
        requirements: {
          scopes: ['mail'],
          capabilities: [{ name: 'Mailbox', minVersion: '1.5' }],
        },
        runtimes: [
          {
            id: 'TaskpaneRuntime',
            type: 'general',
            code: { page: `${baseUrl}/taskpane` },
            lifetime: 'short',
            actions: [
              {
                id: 'openMailLogTaskpane',
                type: 'openPage',
                view: 'MailLogTaskpane',
                pinnable: options.supportsPinning !== false,
              },
            ],
          },
        ],
        ribbons: [
          {
            contexts: ['mailRead', 'mailCompose'],
            tabs: [
              {
                builtInTabId: 'TabDefault',
                groups: [
                  {
                    id: 'MailLogGroup',
                    label: 'MailLog AI',
                    icons: [
                      { size: 16, url: `${baseUrl}/assets/icon-16.png` },
                      { size: 32, url: `${baseUrl}/assets/icon-32.png` },
                      { size: 80, url: `${baseUrl}/assets/icon-80.png` },
                    ],
                    controls: [
                      {
                        id: 'MailLogTaskpaneBtn',
                        type: 'button',
                        label: '메일 분석·일지',
                        icons: [
                          { size: 16, url: `${baseUrl}/assets/icon-16.png` },
                          { size: 32, url: `${baseUrl}/assets/icon-32.png` },
                          { size: 80, url: `${baseUrl}/assets/icon-80.png` },
                        ],
                        supertip: {
                          title: '메일 분석·일지',
                          description:
                            '일자별·기간별 받은 메일, 보낸 메일, 선택한 메일을 분석해 업무 일지와 블로그 글을 자동 생성합니다.',
                        },
                        actionId: 'openMailLogTaskpane',
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };

  return JSON.stringify(manifest, null, 2);
}

/**
 * Generates a Windows Registry (.reg) script for Classic Outlook Shared Folder Trusted Catalog
 * so internal company teams can sideload manifest.xml from a local or SMB network folder.
 */
export function buildWindowsRegistryCatalog(sharedFolderPath: string): string {
  const sanitized = sharedFolderPath.trim() || '\\\\fileserver\\outlook-addins\\maillog';
  const escaped = sanitized.replace(/\\/g, '\\\\');
  return [
    'Windows Registry Editor Version 5.00',
    '',
    '[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\{7A8F2C14-9B5E-4D1C-A3E8-6F4B2D901C52}]',
    '"Id"="{7A8F2C14-9B5E-4D1C-A3E8-6F4B2D901C52}"',
    `"Url"="${escaped}"`,
    '"Flags"=dword:00000001',
    '',
  ].join('\r\n');
}

export function buildCommandsHtml(): string {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <title>MailLog Studio Commands</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script>
    Office.onReady(function () {
      // Outlook Ribbon command runtime ready.
    });
  </script>
</head>
<body></body>
</html>`;
}
