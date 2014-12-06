$ = jQuery.noConflict();

var reportsArray = [];
var url = 'http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/select_action_asset.php';

var report_ids = [1,2,3];

jQuery(document).ready(function() {

    $('div#loading-image').show();

    var current = 0;

    function ajax_call() {
        //check to make sure there are more requests to make
        if (current < report_ids.length) {

            //make the AJAX request with the given data from the `ajaxes` array of objects
            jQuery.ajax({
                type: "GET",
                url: url,
                data: "report_id="+report_ids[current],
                statusCode: {
                    404: function() {
                        alert( "page not found" );
                    },
                    500: function() {
                        alert("internal server error");
                    }
                },
                dataType: "json",
                success  : function (serverResponse) {

                    //stop showing the loading sign until someone uses select
                    $('div#loading-image').hide();
                    //saving the response (object) in the array
                    reportsArray.push(serverResponse);
                    //enabling the <a id="report".report_id."">
                    var rep_id = "a#report"+report_ids[current];
                    $(rep_id).show();
                    current++;
                    ajax_call();
                },
                error: function (serverResponse) {
                    reportsArray.push(0);
                    alert("ERROR: server failure for query: report no."+(current+1)+", "+report_names[current]);
                    var opt = document.createElement("option");
                    var optText = document.createTextNode(report_names[current]);
                    opt.appendChild(optText);
                    opt.setAttribute("value", ""+(current+1)+"");
                    opt.setAttribute("disabled", "1");
                    select_handler.appendChild(opt);

                    current++;
                    ajax_call();
                }
            });
        }
    }

    ajax_call();

    jQuery("a#report_id1").click(function() {
        var remove_el = document.getElementById("report_container");
        if (remove_el != null) {
            remove_el.remove();
        }
        //alert ("Report change event triggered!");
        var report_id = 1;
        //all of the reports are already downloaded

    });

    jQuery("a#report_id2").click(function() {
        var remove_el = document.getElementById("report_container");
        if (remove_el != null) {
            remove_el.remove();
        }
        //alert ("Report change event triggered!");
        var report_id = 2;
        //all of the reports are already downloaded

    });

    jQuery("a#report_id3").click(function() {
        var remove_el = document.getElementById("report_container");
        if (remove_el != null) {
            remove_el.remove();
        }
        //alert ("Report change event triggered!");
        var report_id = 3;
        //all of the reports are already downloaded

    });

});

