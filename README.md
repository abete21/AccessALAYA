# Access ALAYA – Framework for Multi‑User Microsoft Access Applications  
**Version 2.03**  
**Author:** abete21  
**Website:** https://accessalaya.jp  
**Contact:** AccessALAYA@gmail.com  

---

## Revision History
- **2.03 (2026‑06‑14)** – First release of English version. Translated the message and explanatory document.  
- **2.00 (2026‑05‑29)** – Started English version development.  
- **1.20 (2025‑04‑29)** – Added class modules for form controls and composite key support.  
- **1.12 (2024‑03‑30)** – Added SQL loop definitions and external macro commands.  
- **1.10 (2023‑09‑02)** – Added standard form generation feature.  
- **1.09 (2023‑08‑26)** – Introduced installer module.  

---

## Overview
**Access ALAYA** is a framework designed to make Microsoft Access safe and stable for **multi‑user environments**.

Access is a powerful tool for small‑scale database applications, offering the simplicity of Excel with the structure of a relational database.  
However, Access was originally designed for **single‑user or very small teams**, and running it with many users under default settings often leads to file corruption or instability.

Access ALAYA solves these issues and enables **dozens of concurrent users** to safely operate an Access‑based system.

In the author's production environment, ALAYA has been running **100 registered users** with **around 30 concurrent logins**, operating **over a year without maintenance**.

---

## Key Features

### **Multi‑User Ready**
ALAYA is designed from the ground up for multi‑user operation.  
No additional tuning is required later—your system is multi‑user capable from day one.

### **Low‑Code Development**
Common functions are provided as reusable libraries, reducing the amount of VBA code you need to write.  
This improves readability, maintainability, and development speed.

### **Versioning & Deployment Management**
Frequent updates are essential for end‑user applications.  
ALAYA includes built‑in mechanisms for:
- Module updates  
- Version control  
- Automated distribution to end users  

---

## System Requirements
- **Windows 10 or later**  
- **Microsoft Access (32‑bit or 64‑bit)**  
  - Verified on Access 2013 and later  
  - Avoids Windows‑specific system calls for maximum compatibility  
- **Note:** Access is not included in Microsoft Office for Mac, so ALAYA does not run on macOS.

---

## 📦 Included Files (ZIP Package Structure)

The release archive `AccessALAYA2.03.zip` contains the following folders and files:

---

### ⚠️ Important: Startup Script Provided as `.txt`

To prevent GitHub from misclassifying the project as VBScript,  
the startup script is included as:

**`Start ALAYA v3.txt`**

Before using ALAYA, please rename this file to:

**`Start ALAYA v3.vbs`**

This script is required for versioning and automatic deployment.

---

- AccessALAYA/
  - README.md — English version of this document  
  - README_JP.md — Japanese version of this document  
  - docs/
    - Starting AccessALAYA.pptx — Getting started guide  
  - sample/
    - ReadMe.md — Sample application description  
  - src/
    - ALAYA Installer/ — Installer modules  
      - ALAYA Installer1.01.accdb  
      - ALAYA Installer1.02.accdb  
      - ALAYA Installer1.03.accdb — Latest installer version  
    - ALAYADB/ — Sample shared tables  
      - Membership.accdb  
      - SchedIndex.accdb  
      - Schedule.accdb  
      - Tasks.accdb  
    - ALAYAMenu/ — Main GUI module and related files  
      - ALAYA Menu**.accdb — Main user interface module  
      - ALAYA Menu.rel — Release definition file  
      - StartALAYA.vbs — Startup script for versioning and deployment  
      - PrivateDB/ — Local module DB folder (empty in v2.03)  
      - xlTemplate/ — Excel template folder (empty in v2.03)  
    - DBBackup/ — Backup folder (empty for installation)

Each folder corresponds to a component of the Access ALAYA framework:

- **src/** — Core framework modules and installer  
- **sample/** — Demonstration application  
- **docs/** — Documentation and setup guide  
- **README.md / README_JP.md** — English and Japanese overview files  

After extracting the ZIP, open the sample application in Microsoft Access to explore the framework.

---

## Installation
1. Place the ALAYA package on a **file server**.  
2. Run the latest **ALAYA Installer** (`ALAYA Installer**.accdb`) from the installer folder.  
3. The installer will deploy the ALAYA system to the server.

---

## Uninstallation
ALAYA does **not** write to the Windows registry.  
To uninstall, simply delete the deployed project folder from the file server.

---

## How to Use
- Start ALAYA **via the startup script**, not by opening the GUI `.accdb` directly.  
- A `.vbs` startup script is automatically generated in the menu module folder.  
- Distribute **shortcuts to the script** to end users.

---

## License
- Access ALAYA is free for **personal use** and **academic use**.  
- Commercial use requires a license agreement.  
- See the LICENSE file for full terms.

---

## Support & Contact
For questions, please use the official Q&A forum:  
https://accessalaya.jp/alaya-qa/

For other inquiries:  
**AccessALAYA@gmail.com**

---
