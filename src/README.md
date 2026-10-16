# FOX (IDEA NXT) Cryptographic Algorithm - 64-bit VBA Implementation

⚠️ **CRITICAL NOTICE: FOR EDUCATIONAL & CONCEPTUAL PURPOSES ONLY** ⚠️
*This repository is a strict clean-room implementation built to explore the boundaries of the VBA environment. It is NOT intended, licensed, or safe for production deployment or commercial use.*

An educational, conceptual implementation of the **FOX block cipher (FOX-128/16/256)** written entirely in **VBA (Visual Basic for Applications)**. 

This project was developed as an engineering and cryptographic challenge to implement complex primitives within the highly restrictive, non-standard architecture of VBA.

## 🧠 Architectural Constraints & Design

To ensure strict execution linearity and mitigate side-channel vectors within VBA, the code adheres to the following paradigms:
* **Bitwise Domination:** Logic is driven completely via bitwise operations, masking, and `Long` data types coupled with simple addition/subtraction arithmetic.
* **Constant-Time Direction:** Completely avoids conditional branches (`If/Else` for core operations), multiplication, division, and Lookup Tables (LUTs).
* **Pointer Arithmetic in VBA:** Standard VBA string manipulation was discarded. Instead, Windows Kernel API calls (`RtlMoveMemory`) and direct memory pointers (`LongPtr`) are utilized to process input strings of any length.
* **Pure ECB Mode:** Operating strictly in Electronic Codebook (ECB) mode for straightforward validation against official test vectors and educational clarity.

## 💻 Environment & Compatibility

* **Host Application:** MS Office LTSC Pro Plus 2021 (MS Access Container)
* **OS Environment:** Windows 11 Enterprise 25H2
* **Architecture:** Strictly **64-bit VBA** environment (uses 64-bit safe API declarations).
* **Code Page:** Engineered and tested for `HR-ENG` (Croatian/English) code pages.

## 📂 Application Structure & Architecture

The application is modularly organized, separating the graphical user interface from the core cryptographic and pointer-level logic:

### 🖥️ User Interface (Forms)
* **`Fox_128_16_256`** (Main Form): An interactive, bilingual (HR/ENG) control panel. It contains:
  * **2 Input Fields:** `Tekst (Text)` and `Ključ (Key)` for live testing.
  * **5 Command Buttons:** 
    * `Šifriraj (Encrypt)` & `Dešifriraj (Decrypt)` for core operation testing.
    * `Test vektora (Vector Test)` for automated verification against official FOX test vectors.
    * `Briši unos (Erase Input)` & `Izlaz (Quit)`.
  * **Automated UI & Optimization:** On startup, the database automatically initializes the main form. Upon exit, it programmatically executes a *Compact & Repair* routine to ensure the file remains clean and unbloated.
  * **High-Precision Benchmarking:** Dedicates two visual fields to display microsecond-accurate encryption and decryption times powered by `QueryPerformanceCounter` and `QueryPerformanceFrequency`.

### ⚙️ Source Modules (Code)
* **1 Public Module:**
  * **`Funkcije`:** Houses the globally accessible functions and vector values.
* **2 Private Modules:**
  * **`Fox`:** Implements the core mathematical layout of the Fox 128/16/256 cipher (the strict constant-time code logic ( as much as possible ) for round transformations, bitwise masks, and linear arithmetic).
  * **`ObradaTeksta`:** Manages memory pointers, handles character/byte page translations (`HR-ENG`), and wraps the Windows Kernel API (`RtlMoveMemory`) for safe boundary operations.

*Note: Raw code exports (`.bas`, `.cls``) are available in the root directory for immediate reading, while the full demo database is located in the `/src` folder.*

## 👤 About the Author & Motivation

This project is entirely a passion driven by personal curiosity and engineering interest. I am not a professional mathematician or a software engineer by trade, but rather a self-taught technology enthusiast. 

Developing this cipher was a personal milestone in exploring low-level computing concepts, pointer arithmetic, and strict optimization constraints inside an unexpected environment like Microsoft Access. 

## 📚 References & Academic Attribution

This project is a *clean-room implementation* created solely for educational purposes, built in recognition and respect of the groundbreaking cryptographic work by the original authors of the FOX cipher:

* **Original Authors:** **Prof. Serge Vaudenay** and **Dr. Pascal Junod** (developed at **EPFL - École Polytechnique Fédérale de Lausanne**).
* **Reference Specifications:** Based entirely on the publicly available academic document **"FOX Specifications Version 1.2"**.
* **Historical context:** [Wikipedia - IDEA NXT](https://en.wikipedia.org/wiki/IDEA_NXT)

*Legal Note: Inquiries regarding publication permissions were sent to the original authors; however, no response has been received to date. This repository contains completely original source code implementing public mathematical specifications and is not affiliated with, nor endorsed by, the original authors, EPFL, or the original patent holders.*

## 📄 Non-Commercial License
This project is restricted to **Non-Commercial Educational Use Only**. Any commercial application, integration into production software, or monetary exploitation is strictly prohibited. See the accompanying `LICENSE` file for details.
