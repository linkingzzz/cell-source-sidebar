# Installer

This folder will contain installer sources and scripts for Windows (.msi/.exe) and Mac (.pkg / scripts).

Windows installer tasks (MVP):
- Copy Office Add-in static assets to local folder or ensure manifest points to HTTPS URL.
- Register the add-in manifest in user's Office Add-ins local catalog or register per-user side-load entry so Excel shows the add-in.
- Copy WPS plugin files into WPS plugin directory and perform any necessary registration.

Mac installer tasks (MVP):
- Provide scripts to place manifest on a reachable HTTPS endpoint or local path and attempt to register add-in via AppleScript if possible.
- Provide clear guided steps for the user if full automation is not possible.

