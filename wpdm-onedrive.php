<?php
/*
  Plugin Name: WPDM - OneDrive
  Description: OneDrive Explorer for WordPress Download Manager
  Plugin URI: http://www.wpdownloadmanager.com/
  Author: Jesmeen
  Version: 2.0.0
  Author URI: http://www.wpdownloadmanager.com/
 */
namespace WPDM\AddOn;

if (defined('WPDM_VERSION')) {
    require dirname(__FILE__) . '/liveconnect.php';


    if (!defined('WPDM_CLOUD_STORAGE'))
        define('WPDM_CLOUD_STORAGE', 1);

    class OneDrive {

        function __construct() {

            add_action("wpdm_cloud_storage_settings", array($this, "Settings"));
            add_action('wpdm_attach_file_metabox', array($this, 'BrowseButton'));
            add_action('admin_init', array($this, 'wpdm_onedrive_verification'));
        }

        function wpdm_onedrive_verification() {
            $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));
            // Only add hooks when the current user has permissions AND is in Rich Text editor mode
            if (( current_user_can('edit_posts') || current_user_can('edit_pages') ) && get_user_option('rich_editing') && isset($wpdm_onedrive['client_id'])) {

                // Live Connect JavaScript library
                wp_register_script('liveconnect',  'https://js.live.net/v5.0/wl.js');
                wp_enqueue_script('liveconnect');

                wp_enqueue_script('jquery');

                wp_register_script('wpdmonedrive', plugins_url('js/onedrive.js', __FILE__), array('jquery'));
                wp_enqueue_script('wpdmonedrive');
            }
        }

        function Settings() {
            global $current_user;
            if (isset($_POST['__wpdm_onedrive']) && count($_POST['__wpdm_onedrive']) > 0) {
                update_option('__wpdm_onedrive', $_POST['__wpdm_onedrive']);
                die('Settings Saves Successfully!');
            }
            $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));
            ?>
            <div class="panel panel-default">
                <div class="panel-heading"><b><?php _e('OneDrive API Credentials', 'wpdmpro'); ?></b></div>

                <table class="table">



                    <tr>
                        <td>Redirect Url</td>
                        <td><input type="text" name="__wpdm_onedrive[redirect_url]" class="form-control"
                                   value="<?php echo isset($wpdm_onedrive['redirect_url']) ? $wpdm_onedrive['redirect_url'] : ''; ?>"/>
                        </td>
                    </tr>

                    <tr>
                        <td>Client ID</td>
                        <td><input type="text" name="__wpdm_onedrive[client_id]" class="form-control"
                                   value="<?php echo isset($wpdm_onedrive['client_id']) ? $wpdm_onedrive['client_id'] : ''; ?>"/>
                        </td>
                    </tr>

                    <tr>
                        <td>Client Secret</td>
                        <td><input type="text" name="__wpdm_onedrive[client_secret]" class="form-control"
                                   value="<?php echo isset($wpdm_onedrive['client_secret']) ? $wpdm_onedrive['client_secret'] : ''; ?>"/>
                        </td>

                    <tr>
                        <td>Tenant ID</td>
                        <td><input type="text" name="__wpdm_onedrive[tenant_id]" class="form-control"
                                   value="<?php echo isset($wpdm_onedrive['tenant_id']) ? $wpdm_onedrive['tenant_id'] : ''; ?>"/>
                        </td>
                    </tr>

                    <tr>
                        <td>Base Url</td>
                        <td><input type="text" name="__wpdm_onedrive[base_Url]" class="form-control"
                                   value="<?php echo isset($wpdm_onedrive['base_Url']) ? $wpdm_onedrive['base_Url'] : ''; ?>"/>
                        </td>
                    </tr>
                </table>

            </div>


            <?php
        }

        function BrowseButton() {
            $wpdm_onedrive = maybe_unserialize(get_option('__wpdm_onedrive', array()));
            ?>
            <div class="w3eden">

                <script type="text/javascript" src="https://alcdn.msauth.net/browser/2.19.0/js/msal-browser.min.js"></script>
                <!-- <script type="text/javascript" src="<?= home_url('/wp-content/plugins') ?>/wpdm-onedrive/js/auth.js"></script> -->
                <script type="text/javascript">

                    const baseUrl = "<?php echo wpdm_valueof($wpdm_onedrive, 'base_Url'); ?>"
                    // const baseUrl = "https://iyulad-my.sharepoint.com/";
                    //const baseUrl = "https://onedrive.live.com/";
                    function combine(...paths) {

                        return paths
                            .map(path => path.replace(/^[\\|/]/, "").replace(/[\\|/]$/, ""))
                            .join("/")
                            .replace(/\\/g, "/");
                    }


                    const params = {
                        sdk: "8.0",
                        entry: {
                            oneDrive: {
                                files: {},
                            }
                        },
                        authentication: {},
                        messaging: {
                            origin: "<?= home_url() ?>",
                            channelId: "27"
                        },
                        typesAndSources: {
                            mode: "all",
                            pivots: {
                                oneDrive: true,
                                recent: true,
                            },
                        },
                    };

                    const msalParams = {
                        auth: {
                            authority: "https://login.microsoftonline.com/<?php echo wpdm_valueof($wpdm_onedrive, 'tenant_id'); ?>/",
                            clientId: "<?php echo wpdm_valueof($wpdm_onedrive, 'client_id'); ?>",
                            redirectUri: "<?php echo wpdm_valueof($wpdm_onedrive, 'redirect_url'); ?>"
                        },
                    }

                    const app = new msal.PublicClientApplication(msalParams);

                    async function getToken(command) {
                        let accessToken = "";
                        let authParams = null;

                        switch (command.type) {
                            case "SharePoint":
                            case "SharePoint_SelfIssued":
                                authParams = { scopes: [`${combine(command.resource, ".default")}`] };
                                break;
                            default:
                                break;
                        }

                        try {
                            const resp = await app.acquireTokenSilent(authParams);
                            accessToken = resp.accessToken;
                        } catch (e) {
                            const resp = await app.loginPopup(authParams);
                            app.setActiveAccount(resp.account);

                            if (resp.idToken) {
                                const resp2 = await app.acquireTokenSilent(authParams);
                                accessToken = resp2.accessToken;
                            }
                        }

                        return accessToken;
                    }
                    let win = null;
                    let port = null;


                    async function launchOneDrivePicker() {
                        // var pickerOptions = {
                        // clientId: "<?= $wpdm_onedrive['client_id'] ?>",
                        //     action: "download",
                        //     advanced: {
                        //         redirectUri: "https://wpdm.pro/wp-admin/post-new.php?post_type=wpdmpro"
                        //     },
                        //     success: function (files) {
                        //         // Handle returned file object(s)
                        //         var id = files.values[0].fileName.replace(/([^a-zA-Z0-9]*)/g, "");
                        //         InsertOneDriveLink(files.values[0].link, id, files.values[0].fileName);
                        //     },
                        //     cancel: function () {
                        //         alert("You picked failed");
                        //         // handle when the user cancels picking a file
                        //     },
                        //     error: function(error) { console.log(error); },
                        //     linkType: "webViewLink", // or "downloadLink",
                        //     multiSelect: false // or true
                        // };
                        // OneDrive.open(pickerOptions);

                        win = window.open("", "Picker", "width=800,height=600")

                        const authToken = await getToken({
                            resource: baseUrl,
                            command: "authenticate",
                            type: "SharePoint",
                        });

                        const queryString = new URLSearchParams({
                            filePicker: JSON.stringify(params),
                        });

                        const url = combine(baseUrl, `_layouts/15/FilePicker.aspx?${queryString}`);

                        const form = win.document.createElement("form");
                        form.setAttribute("action", url);
                        form.setAttribute("method", "POST");
                        win.document.body.append(form);

                        const input = win.document.createElement("input");
                        input.setAttribute("type", "hidden")
                        input.setAttribute("name", "access_token");
                        input.setAttribute("value", authToken);
                        form.appendChild(input);

                        form.submit();


                    window.addEventListener("message", async (event) => {
                        if (event.source && event.source === win) {
                            const message = event.data;

                            if (message.type === "initialize" && message.channelId === params.messaging.channelId) {
                                port = event.ports[0];
                                port.addEventListener("message", messageListener);
                                port.start();
                                port.postMessage({
                                    type: "activate",
                                });
                            }
                        }
                    });

                    }

                    // function InsertOneDriveLink(file, id, name) {

                    //     <?php if (version_compare(WPDM_VERSION, '4.0.0', '>')) { ?>
                    //         var html = jQuery('#wpdm-file-entry').html();
                    //         var ext = 'png'; //response.split('.');
                    //         //ext = ext[ext.length-1];
                    //         name = file.substring(0, 80) + "...";
                    //         var icon = "<?php echo WPDM_BASE_URL; ?>file-type-icons/48x48/" + ext + ".png";
                    //         html = html.replace(/##filepath##/g, file);
                    //         //html = html.replace(/##filepath##/g, file);
                    //         html = html.replace(/##fileindex##/g, id);
                    //         html = html.replace(/##preview##/g, icon);
                    //         jQuery('#currentfiles').prepend(html);

                    //     <?php } else { ?>
                    //         jQuery('#wpdmfile').val(file + "#" + name);
                    //         jQuery('#cfl').html('<div><strong>' + name + '</strong>').slideDown();
                    //     <?php } ?>
                    // }

                    async function createPublicDownloadLink(itemId) {
                        const accessToken = await getToken({
                            resource: baseUrl,
                            command: "authenticate",
                            type: "SharePoint",
                        });
                        console.log(accessToken);


                        var myHeaders = new Headers();
                        myHeaders.append("Authorization", "Bearer "+accessToken);

                        var requestOptions = {
                            method: 'POST',
                            headers: myHeaders,
                            redirect: 'follow'
                        };

                        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/createLink`;

                        try {

                            const response = await fetch(url, {
                                method: 'POST',
                                headers: {
                                    'Authorization': `Bearer ${accessToken}`,
                                    'Content-Type': 'application/json'
                                },
                                body: JSON.stringify({
                                    type: 'view',
                                    scope: 'anonymous',
                                    retainInheritedPermissions: false,

                                })
                            });
                            // const data = await response.json();

                            console.log('DL', response);
                            // return data.link.webUrl;

                            const data = await response.json();
                            console.log(data);

                            if (data.link && data.link.webUrl) {
                                console.log(`Public Download Link: ${data.link.webUrl}`);
                                return data.link.webUrl;
                            } else {
                                console.error('Error: Unexpected response format:', data);
                                return null;
                            }
                        } catch (error) {
                            console.error('Error creating public download link:', error);
                            return null;
                        }
                    }

                    async function messageListener(message) {
                        switch (message.data.type) {

                            case "notification":
                                    console.log(`notification: ${message.data}`);
                                    break;
                            case "command":
                                port.postMessage({
                                    type: "acknowledge",
                                    id: message.data.id,
                                });

                                const command = message.data.data;

                                switch (command.command) {

                                    case "authenticate":


                                            const token = await getToken(command);

                                            if (typeof token !== "undefined" && token !== null) {

                                                port.postMessage({
                                                    type: "result",
                                                    id: message.data.id,
                                                    data: {
                                                        result: "token",
                                                        token,
                                                    }
                                                });

                                            } else {
                                                console.error(`Could not get auth token for command: ${JSON.stringify(command)}`);
                                            }

                                            break;

                                        case "close":

                                            win.close();
                                            break;


                                    case "pick":
                                        console.log(`Picked: ${JSON.stringify(command)}`);

                                        const pickedItem = command.items[0];

                                        //shareable link part
                                        // if (pickedItem) {
                                        //     const publicDownloadLink = await createPublicDownloadLink(pickedItem.id);
                                        //     console.log('Public Download Link:', publicDownloadLink);

                                        //     document.getElementById("downloadButton").style.display = "block";
                                        //     document.getElementById("downloadButton").addEventListener("click", async () => {
                                        //     window.open(publicDownloadLink, "_blank");
                                        //     });
                                        // }

                                        //document.getElementById("pickedFiles").innerHTML = `<pre>${JSON.stringify(command, null, 2)}</pre>`;
                                        console.log(JSON.stringify(command, null, 2));

                                        port.postMessage({
                                            type: "result",
                                            id: message.data.id,
                                            data: {
                                                result: "success",
                                            },
                                        });

                                        win.close();

                                        let publicDownloadLink = await createPublicDownloadLink(pickedItem.id);
                                        console.log(publicDownloadLink);

                                        break;


                                    default:

                                    console.warn(`Unsupported command: ${JSON.stringify(command)}`, 2);

                                    port.postMessage({
                                        result: "error",
                                        error: {
                                            code: "unsupportedCommand",
                                            message: command.command
                                        },
                                        isExpected: true,
                                    });
                                    break;

                                }
                                break;
                        }
                    }



                </script>


                <a href="#" id="btn-onedrive" style="margin-top: 10px;" title="OneDrive" onclick="return launchOneDrivePicker()" class="btn wpdm-onedrive btn-block">Select From OneDrive</a>

            </div>


            <?php
        }

    }

    new OneDrive();
}

