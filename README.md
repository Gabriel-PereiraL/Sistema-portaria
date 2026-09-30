# 🎟️ Museum Entry Management System

A real-world system developed for the **Museu da Imaginação**, currently in active use for managing visitor entry and access.

---

## 📖 Overview

This system was designed to handle the museum's entry workflow with a focus on:

* Security (fully offline operation)
* Simplicity and usability
* Reliability for continuous daily use

---

## ⚙️ Features

* Local-first architecture (no internet required)
* Intuitive interface for staff usage
* Visitor management and control system
* Reliable operation for kiosk/desk environments

---

## 🛠️ Tech Stack

* **Python (PyQt)** — interface
* **MariaDB** — database

---

## 🧪 Environment

* Desktop application
* Designed for internal institutional use

---

## 🚀 Status

✔ Currently deployed and actively used in production at the museum.

For existing installations, I apply [`schema_hardening.sql`](schema_hardening.sql) once after resolving any pre-existing duplicate open visits. The generated column and unique index enforce one open entry per visitor in MariaDB; the application also locks the visitor row and performs the check and insert in one transaction.

---

## 👤 Author

Gabriel Vitor
