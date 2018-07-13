# AVAYA-CMS-Monitor

System try tool to monitor the Calls status on AVAYA systems

The tool is launching the CMS script (save to HTML script) and read the HTML generated.

On the system bar the icon will show the agents status and the availability along with the calls waiting (if there are).

If no agents are available or calls are waiting a baloon message will be prompted.

Change cmsmonitor.xml settings for your needs.

<server> tag must be configured with the user name of the machine that will act as server.

Config structure:

\<config>\n
\<script_path>Save HTML-QueueAgent Status SL.acsauto\</script_path>  // Path to the htlm generator script \n
\<start>08:00 AM\</start> // Starting time for monitoring \n
\<stop>07:00 PM\</stop> // Ending time for monitoring \n
\<server>xxxx\</server> // User name server \n
\</config>\n



