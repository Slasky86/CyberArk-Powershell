# How to use Remote Desktop Manager from Devolutions free version towards Cyberark PSM

1. Install RDM from Devolutions
2. Start RDM and log in (either with a free account or an enterprise account)
3. Import the template from this repo
4. Create a folder if you want to group connections
5. Add entry -> Add from template
6. Select the template you imported in step 3 and replace information thats within brackets <>, including the brackets

The target account is defined as username@address, where the fields corrosponds with the fields in CyberArk.

## What does this template do

This template pre-fills some information and settings. It sets that any special key/shortcut pressed is sent to the remote machine (for instance pressing the Windows button will make it react in the remote server rather than your local client). The template also deactivates NLA (which is required for a connection to the PSM to work). 

The template does not define any credentials as you have to provide these yourself. These can be set at a folder level and the connections can be set to inheret them from the top folder. This is the credentials used to authenticate yourself and verify that you have access to the given credentials you are trying to retrieve. If you don't have access to the credentials, the connection won't work.


## Requirements

This requires that the vault can authorize your logon request, which means the user has to be a transparent user or a vault user of which you know the password. If its a transparent user, remember to insert the username as it displays in the vault.

# Using Remote Desktop Manager from Devolutions licensed version

If you pay for a license for RDM you can create your own CyberArk dashboard (described here):
https://kb.devolutions.net/kb_rdm_cyberark_dashboard_configuration.html

https://blog.devolutions.net/2020/10/going-passwordless-with-remote-desktop-manager-and-cyberark/

This allows you to log on to the PVWA with SAML and retrieve your available accounts. This is similar to the PSMClient that CyberArk offers. 
