//
//  DiscordRPC.swift
//  MSOffice DiscordRPC
//
//  Created by Oryon Ivanov on 30/08/22.
//

import Foundation
import SwordRPC
import PythonKit
import Foundation
import AppKit

//These functions find's where the application is stored and if its running
func appget_word() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft Word.app/Contents/MacOS/Microsoft Word") {
            print("Located Word")
            return true
        } else {
            print("Word not found (You sure you located it properly?)")
            return false
        }
    }

func appget_excel() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft Excel.app/Contents/MacOS/Microsoft Excel") {
            print("Located Excel")
            return true
        } else {
            print("Excel not found (You sure you located it properly?)")
            return false
        }
    }

func appget_powerpoint() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft PowerPoint.app/Contents/MacOS/Microsoft PowerPoint") {
            print("Located PowerPoint")
            return true
        } else {
            print("PowerPoint not found (You sure you located it properly?)")
            return false
        }
    }

func appget_onenote() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft OneNote.app/Contents/MacOS/Microsoft OneNote") {
            print("Located OneNote")
            return true
        } else {
            print("OneNote not found (You sure you located it properly?)")
            return false
        }
    }

func appget_outlook() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft Outlook.app/Contents/MacOS/Microsoft Outlook") {
            print("Located Outlook")
            return true
        } else {
            print("Outlook not found (You sure you located it properly?)")
            return false
        }
    }

func appget_teams() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/Microsoft Teams.app/Contents/MacOS/Teams") {
            print("Located Teams")
            return true
        } else {
            print("Teams not found (You sure you located it properly?)")
            return false
        }
    }

func appget_onedrive() -> Bool {
        if FileManager.default.fileExists(atPath: "/Applications/OneDrive.app/Contents/MacOS/OneDrive") {
            print("Located OneDrive")
            return true
        } else {
            print("OneDrive not found (You sure you located it properly?)")
            return false
        }
    }
