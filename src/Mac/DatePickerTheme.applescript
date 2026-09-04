on GetDarkMode(paramString)
    tell application "System Events"
        tell appearance preferences
            return dark mode
        end tell
    end tell
end GetDarkMode