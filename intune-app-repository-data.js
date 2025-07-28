const applications = [
  // Microsoft Products
  {
    name: ".NET Framework",
    publisher: "Microsoft Corporation",
    category: "Runtime",
    msiUrl: "https://dotnet.microsoft.com/download/dotnet-framework",
    armUrl: "https://dotnet.microsoft.com/download/dotnet/8.0",
  },
  {
    name: ".NET Runtime 8",
    publisher: "Microsoft Corporation",
    category: "Runtime",
    msiUrl: "https://dotnet.microsoft.com/download/dotnet/8.0",
    armUrl: "https://dotnet.microsoft.com/download/dotnet/8.0",
  },
  {
    name: "Microsoft Edge",
    publisher: "Microsoft",
    category: "Browser",
    msiUrl:
      "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/5838c4cb-c4e1-40ec-af9a-0a5f8cd78e24/MicrosoftEdgeEnterpriseX64.msi",
    armUrl:
      "https://msedge.sf.dl.delivery.mp.microsoft.com/filestreamingservice/files/2dbbce27-c976-4f7c-b493-734b3262e2d2/MicrosoftEdgeEnterpriseARM64.msi",
  },
  {
    name: "Microsoft Teams",
    publisher: "Microsoft",
    category: "Communication",
    msiUrl: "https://www.microsoft.com/microsoft-teams/download-app",
    armUrl: "https://www.microsoft.com/microsoft-teams/download-app",
  },
  {
    name: "Microsoft OneDrive",
    publisher: "Microsoft",
    category: "Cloud Storage",
    msiUrl: "https://www.microsoft.com/microsoft-365/onedrive/download",
    armUrl: "https://www.microsoft.com/microsoft-365/onedrive/download",
  },
  {
    name: "Microsoft Visual Studio Code",
    publisher: "Microsoft Corporation",
    category: "Development",
    msiUrl: "https://code.visualstudio.com/download",
    armUrl: "https://code.visualstudio.com/download",
  },
  {
    name: "Microsoft PowerToys",
    publisher: "Microsoft Corporation",
    category: "Utilities",
    msiUrl: "https://github.com/microsoft/PowerToys/releases",
    armUrl: "https://github.com/microsoft/PowerToys/releases",
  },
  {
    name: "Microsoft Power BI Desktop",
    publisher: "Microsoft",
    category: "Analytics",
    msiUrl:
      "https://download.microsoft.com/download/8/8/0/880bca75-79dd-466a-927d-1abf1f5454b0/PBIDesktopSetup_x64.exe",
  },

  // Adobe Products
  {
    name: "Adobe Acrobat Reader DC",
    publisher: "Adobe Inc.",
    category: "PDF",
    msiUrl: "https://get.adobe.com/reader/enterprise/",
  },
  {
    name: "Adobe Creative Cloud",
    publisher: "Adobe Inc.",
    category: "Creative",
    msiUrl: "https://www.adobe.com/creativecloud/business/teams.html",
  },

  // Security Software
  {
    name: "Avast Antivirus",
    publisher: "Avast",
    category: "Security",
    msiUrl: "https://www.avast.com/business/downloads",
  },
  {
    name: "ESET Endpoint Security",
    publisher: "ESET, spol. s r.o.",
    category: "Security",
    msiUrl: "https://www.eset.com/us/business/download/endpoint-security/",
  },
  {
    name: "Malwarebytes",
    publisher: "Malwarebytes",
    category: "Security",
    msiUrl: "https://www.malwarebytes.com/business/downloads",
  },
  {
    name: "Bitwarden",
    publisher: "Bitwarden Inc.",
    category: "Security",
    msiUrl: "https://bitwarden.com/download/",
  },

  // Browsers
  {
    name: "Google Chrome",
    publisher: "Google",
    category: "Browser",
    msiUrl:
      "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7BBD050720-A9FA-FA92-D25C-940523BD2DC1%7D%26browser%3D0%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dtrue%26ap%3Dx64-stable-statsdef_0%26brand%3DGCEA/dl/chrome/install/googlechromestandaloneenterprise64.msi",
    armUrl:
      "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7BBD050720-A9FA-FA92-D25C-940523BD2DC1%7D%26browser%3D0%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dtrue%26ap%3Darm64-stable-statsdef_0%26brand%3DGCEA/dl/chrome/install/GoogleChromeStandaloneEnterprise_Arm64.msi",
  },
  {
    name: "Mozilla Firefox",
    publisher: "Mozilla",
    category: "Browser",
    msiUrl:
      "https://download-installer.cdn.mozilla.net/pub/firefox/releases/140.0.4/win64/en-US/Firefox%20Setup%20140.0.4.msi",
    armUrl:
      "https://cdn.stubdownloader.services.mozilla.com/builds/firefox-latest-ssl/en-US/win64-aarch64/9e4991bc3be5cb5cc428be35bc7af3dd8156c3a22550d356c043589255e76e03/Firefox%20Setup%20140.0.4.exe",
  },
  {
    name: "Mozilla Firefox ESR",
    publisher: "Mozilla",
    category: "Browser",
    msiUrl: "https://www.mozilla.org/firefox/enterprise/",
    armUrl: "https://www.mozilla.org/firefox/enterprise/",
  },
  {
    name: "Opera",
    publisher: "Opera Software",
    category: "Browser",
    msiUrl:
      "https://download5.operacdn.com/ftp/pub/opera/desktop/120.0.5543.38/win/Opera_120.0.5543.38_Setup_x64.exe",
  },

  // Development Tools
  {
    name: "Git",
    publisher: "The Git Development Community",
    category: "Development",
    msiUrl: "https://git-scm.com/download/win",
    armUrl: "https://git-scm.com/download/win",
  },
  {
    name: "Node.js LTS",
    publisher: "Node.js Foundation",
    category: "Development",
    msiUrl: "https://nodejs.org/en/download/",
    armUrl: "https://nodejs.org/en/download/",
  },
  {
    name: "Notepad++",
    publisher: "Notepad++",
    category: "Development",
    msiUrl: "https://notepad-plus-plus.org/downloads/",
    armUrl: "https://notepad-plus-plus.org/downloads/",
  },
  {
    name: "MySQL Workbench CE",
    publisher: "Oracle",
    category: "Development",
    msiUrl: "https://dev.mysql.com/downloads/workbench/",
  },
  {
    name: "pgAdmin 4",
    publisher: "The pgAdmin Development Team",
    category: "Development",
    msiUrl: "https://www.pgadmin.org/download/pgadmin-4-windows/",
  },
  {
    name: "DBeaver Community",
    publisher: "DBeaver Corp",
    category: "Development",
    msiUrl: "https://dbeaver.io/download/",
  },

  // Utilities
  {
    name: "7-Zip",
    publisher: "Igor Pavlov",
    category: "Utilities",
    msiUrl: "https://www.7-zip.org/download.html",
    armUrl: "https://www.7-zip.org/download.html",
  },
  {
    name: "CCleaner",
    publisher: "Piriform",
    category: "Utilities",
    msiUrl: "https://www.ccleaner.com/ccleaner/download",
  },
  {
    name: "WinRAR",
    publisher: "WinRAR GmbH (Ltd.)",
    category: "Utilities",
    msiUrl: "https://www.win-rar.com/download.html",
  },
  {
    name: "WinSCP",
    publisher: "Martin Prikryl",
    category: "Utilities",
    msiUrl: "https://winscp.net/eng/download.php",
  },
  {
    name: "PuTTY",
    publisher: "PuTTY",
    category: "Utilities",
    msiUrl: "https://www.putty.org/",
    armUrl: "https://www.putty.org/",
  },
  {
    name: "FileZilla Client",
    publisher: "FileZilla",
    category: "Utilities",
    msiUrl: "https://filezilla-project.org/download.php?type=client",
  },

  // Communication
  {
    name: "Slack",
    publisher: "Slack Technologies, Inc.",
    category: "Communication",
    msiUrl:
      "https://downloads.slack-edge.com/desktop-releases/windows/x64/4.44.65/Slack.msix",
    armUrl:
      "https://downloads.slack-edge.com/desktop-releases/windows/arm64/4.44.65/Slack.msix",
  },
  {
    name: "Zoom",
    publisher: "Zoom Video Communications",
    category: "Communication",
    msiUrl: "https://cdn.zoom.us/prod/6.5.3.7509/x64/ZoomInstallerFull.msi",
    armUrl: "https://cdn.zoom.us/prod/6.5.3.7509/arm64/ZoomInstallerFull.msi",
  },
  {
    name: "Cisco Webex Meetings",
    publisher: "Cisco Webex LLC",
    category: "Communication",
    msiUrl:
      "https://binaries.webex.com/WebexOfclDesktop-Win-64-Gold/Webex.msi?_gl=1*1bq9fns*_gcl_au*MTc4MjEzODgzLjE3NTIwMTc3ODE.",
  },
  {
    name: "Cisco Jabber",
    publisher: "Cisco",
    category: "Communication",
    msiUrl:
      "https://www.cisco.com/c/en/us/products/unified-communications/jabber/index.html",
  },
  {
    name: "Skype",
    publisher: "Microsoft",
    category: "Communication",
    msiUrl: "https://www.skype.com/en/business/",
  },

  // Media
  {
    name: "VLC Media Player",
    publisher: "VideoLAN",
    category: "Media",
    msiUrl: "https://www.videolan.org/vlc/download-windows.html",
    armUrl: "https://www.videolan.org/vlc/download-windows.html",
  },
  {
    name: "Audacity",
    publisher: "Audacity Team",
    category: "Media",
    msiUrl: "https://www.audacityteam.org/download/",
  },
  {
    name: "OBS Studio",
    publisher: "OBS Project",
    category: "Media",
    msiUrl: "https://obsproject.com/download",
  },
  {
    name: "HandBrake",
    publisher: "GPL v2",
    category: "Media",
    msiUrl: "https://handbrake.fr/downloads.php",
  },
  {
    name: "GIMP",
    publisher: "The Gimp Team",
    category: "Media",
    msiUrl: "https://www.gimp.org/downloads/",
  },
  {
    name: "Blender",
    publisher: "Blender Foundation",
    category: "Media",
    msiUrl: "https://www.blender.org/download/",
  },
  {
    name: "Paint.NET",
    publisher: "dotPDN LLC",
    category: "Media",
    msiUrl: "https://www.getpaint.net/download.html",
  },

  // Remote Access
  {
    name: "TeamViewer",
    publisher: "TeamViewer",
    category: "Remote Access",
    msiUrl: "https://www.teamviewer.com/en/download/windows/",
    armUrl: "https://www.teamviewer.com/en/download/windows/",
  },
  {
    name: "Splashtop Streamer",
    publisher: "Splashtop Inc",
    category: "Remote Access",
    msiUrl: "https://www.splashtop.com/downloads",
  },
  {
    name: "Citrix Workspace",
    publisher: "Citrix Systems",
    category: "Remote Access",
    msiUrl: "https://www.citrix.com/downloads/workspace-app/",
  },
  {
    name: "VMware Horizon Client",
    publisher: "VMware",
    category: "Remote Access",
    msiUrl: "https://customerconnect.vmware.com/downloads/#all_products",
    armUrl: "https://customerconnect.vmware.com/downloads/#all_products",
  },

  // Cloud Storage
  {
    name: "Dropbox",
    publisher: "Dropbox, Inc",
    category: "Cloud Storage",
    msiUrl: "https://www.dropbox.com/install",
  },
  {
    name: "Google Drive",
    publisher: "Google",
    category: "Cloud Storage",
    msiUrl: "https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe",
  },

  // Office & Productivity
  {
    name: "LibreOffice",
    publisher: "LibreOffice",
    category: "Office",
    msiUrl: "https://www.libreoffice.org/download/download/",
    armUrl: "https://www.libreoffice.org/download/download/",
  },
  {
    name: "Foxit PDF Reader",
    publisher: "Foxit Software",
    category: "PDF",
    msiUrl: "https://www.foxit.com/downloads/",
  },
  {
    name: "PDF-XChange Editor",
    publisher: "PDF-XChange Co Ltd.",
    category: "PDF",
    msiUrl: "https://www.pdf-xchange.com/product/pdf-xchange-editor",
  },
  {
    name: "PDFCreator Free",
    publisher: "pdfforge GmbH",
    category: "PDF",
    msiUrl: "https://www.pdfforge.org/pdfcreator/download",
  },

  // Password Managers
  {
    name: "1Password",
    publisher: "Agilebits Inc.",
    category: "Security",
    msiUrl: "https://downloads.1password.com/win/1PasswordSetup-latest.msi",
    armUrl: "https://1password.com/downloads",
  },
  {
    name: "KeePass Password Safe",
    publisher: "Dominik Reichl",
    category: "Security",
    msiUrl: "https://keepass.info/download.html",
  },

  // Other Tools
  {
    name: "Java",
    publisher: "Oracle",
    category: "Runtime",
    msiUrl: "https://www.java.com/en/download/manual.jsp",
  },
  {
    name: "Wireshark",
    publisher: "GNU General Public License",
    category: "Network",
    msiUrl: "https://www.wireshark.org/download.html",
    armUrl: "https://www.wireshark.org/download.html",
  },
  {
    name: "CrystalDiskInfo",
    publisher: "Crystal Dew World",
    category: "Utilities",
    msiUrl: "https://crystalmark.info/en/software/crystaldiskinfo/",
  },
  {
    name: "Greenshot",
    publisher: "GNU General Public License",
    category: "Utilities",
    msiUrl: "https://getgreenshot.org/downloads/",
  },
  {
    name: "ShareX",
    publisher: "ShareX Team",
    category: "Utilities",
    msiUrl: "https://getsharex.com/downloads",
  },
  {
    name: "IrfanView",
    publisher: "Irfan Skiljan",
    category: "Media",
    msiUrl: "https://www.irfanview.com/64bit.htm",
  },
  {
    name: "XnView MP",
    publisher: "Gougelet Pierre-e",
    category: "Media",
    msiUrl: "https://www.xnview.com/en/xnviewmp/#downloads",
  },
  {
    name: "WinMerge",
    publisher: "Thingamahoochie Software",
    category: "Development",
    msiUrl: "https://winmerge.org/downloads/",
  },
  {
    name: "Far Manager",
    publisher: "Far Group",
    category: "Utilities",
    msiUrl: "https://www.farmanager.com/download.php",
  },
  {
    name: "PeaZip",
    publisher: "Giorgio Tani",
    category: "Utilities",
    msiUrl: "https://peazip.github.io/",
  },
  {
    name: "Cyberduck",
    publisher: "iterate GmbH",
    category: "Utilities",
    msiUrl: "https://cyberduck.io/download/",
  },
  {
    name: "Mozilla Thunderbird",
    publisher: "Mozilla",
    category: "Communication",
    msiUrl: "https://www.thunderbird.net/en-US/thunderbird/all/",
    armUrl: "https://www.thunderbird.net/en-US/thunderbird/all/",
  },
  {
    name: "OpenVPN",
    publisher: "OpenVPN",
    category: "Network",
    msiUrl: "https://openvpn.net/community-downloads/",
  },
  {
    name: "NordLayer VPN",
    publisher: "NordLayer",
    category: "Network",
    msiUrl: "https://downloads.nordlayer.com/win/releases/rss.xml",
  },
  {
    name: "FortiClient VPN",
    publisher: "Fortinet",
    category: "Network",
    msiUrl: "https://www.fortinet.com/support/product-downloads",
  },
  {
    name: "TightVNC",
    publisher: "TightVNC Software",
    category: "Remote Access",
    msiUrl: "https://www.tightvnc.com/download.php",
  },
  {
    name: "Imageglass",
    publisher: "Imageglass",
    category: "Media",
    msiUrl: "https://imageglass.org/download",
  },
  {
    name: "WinZip",
    publisher: "Corel Corporation",
    category: "Utilities",
    msiUrl: "https://www.winzip.com/en/download/",
  },
];
