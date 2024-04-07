# Conan Exiles Server on Windows Server Core

A Powershell script to install everything needed to create and set up a dedicated Conan Exiles server on Windows Server Core.

While running a dedicated game server is best ran on a Linux server, not all games offer this.  This script was written to solve the problem of running a dedicated game server on Windows.  The reason we all love Linux game servers is because they're headless.  Utilizing only the command line to interact with the server, this gives nearly all of the system resources to run the game server app.  Windows generally does not install without a GUI (Graphical User Interface).  However, there is a headless Windows option; Windows Server Core.

Running a Windows Server Core OS can be difficult if you're not real great with Powershell, but is definitely the best way to run a Windows server in this setting.  So, this script fixes all that.  No more will you need to spend hours (or even days) looking up Powershell commands and play the guessing game on if something will work the way you want it to.  It will install everything you need to install and host your very own Conan Exiles Dedicated Server.  I wanted to make this as simple as possible for anyone to use, but if you run into any issues or need any additional help, you can find me on Discord under the same name.

It is highly recommended that you fully update your Windows Server Core OS prior to running the script to ensure proper functionality of the script, but the script will check and apply all Windows updates before doing anything else.  If you are not familiar with Windows Server Core, Option 6 from SConfig will open the Windows Update menu, and Option 1 from the Windows Update menu will check for all updates.  During the script's install process, you will be prompted to enter where you want to install the server, the name you want to use for your server, the password you want to use, and the port numbers you want to use (if you dont want to use the defaults).  Enter in this information and the server's config file will be created for you.  Be sure to reboot the server after everything is complete to ensure any additional updates are fully installed.  Use Port Forwarding to open these ports on your router and send connections to the server's IP address.

To update your existing server, run the UpdateServer.ps1 script.  This will prompt you for the directory that you installed the server, and will auto-update the server.  If you enter the wrong directory, it will error out and prompt you again.

I hope this helps everyone get their own server up and running.  Enjoy!

-TripodGG