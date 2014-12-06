/**
 * Created by Nelia on 7/6/2014.
 */


Number.prototype.format = function(n, x) {
    var re = '\\d(?=(\\d{' + (x || 3) + '})+' + (n > 0 ? '\\.' : '$') + ')';
    return this.toFixed(Math.max(0, ~~n)).replace(new RegExp(re, 'g'), '$&,');
};

function populateIframe(id,path)
{
    var ifrm = document.getElementById(id);
    ifrm.src = "http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/download.php?path="+path+".xlsx";
}

/***************
 * Arrays of "Buttons" to create them and events dynamically
 ***************/
var LArray = [
    [
    "Client/Cost Centre", //0
    "Client/Country", //1 "Revenue_Client_Country"
    "Cost Centre/Country", //2
    "Operator/Country", //3 "Revenue_Operator_Country"
    "Revenue by Service", //4
    "Service/Country", //5 "Revenue_Service_Category_Country"
    "Revenue by Technology", //6
    "Service/Client", //7
    "Service/Country", //8
    "Service/Operator", //9
    "Technology/Country", //10 "Revenue_Technology_Country"
    "Actual/Target/Country", //11
    "Actual/Target for Client/Country", //12 "Revenue_Client_Country_Target"
    "Per Job/Country", //13 "Avg_Revenue_by_Country
    "Per Job/Client", //14
    "Per Job/Operator", //15
    "Per Job/Technology" //16
    ],
   [
    "Assets (general)", //0
    "Assets (summary/value)", //1
    "Assets (summary/quantity)", //2
    "Utilization by Jobs/Tool", //3
    "Utilization by Revenue/Tool", //4
    "Asset Utilization / Tool Turns" //5
    ],
    [
    "Vacation Balance", //0
    "HR Loading"//, //1
    //"New Vacation Balance"
    ]
];

var File_Names = [
    [
        "",
        "Revenue_Client_Country",
        "",
        "Revenue_Operator_Country",
        "",
        "Revenue_Service_Category_Country",
        "",
        "",
        "",
        "",
        "Revenue_Technology_Country",
        "",
        "Revenue_Client_Country_Target",
        "Avg_Revenue_by_Country",
        "",
        "",
        ""
    ],
    [
        "assets1", //0
        "assets_summary", //1
        "", //2
        "", //3
        "", //4
        "Cons_Asset_Util" //5
    ],
    [
        "hr_vacation_balance", //0
        "hr_loading"//, //1
        //"hr_vacation_balance_v2"
    ]
];

var cat_controllers = ["rev", "ast", "hrr"];

var Indexes = [0, LArray[0].length, LArray[0].length + LArray[1].length];

var level1 = '';
var level2 = '';
var level3 = '';

var url = 'http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/read_action.php';

$ = jQuery.noConflict();




jQuery(document).ready(function() {

    for (var cats = 0; cats < cat_controllers.length; cats++) {
        process_controllers(cat_controllers[cats], LArray[cats], File_Names[cats], Indexes[cats]);
    }

});



function process_controllers(controller, links_array, filenames, ind) {
    jQuery("a#"+controller).click(function (e) {
        e.preventDefault();
        var head = document.getElementById("head");
        var report_body = document.getElementById("reports_body");

        var remove_el = document.getElementById("cur_reports");

        if (remove_el != null) {
            report_body.removeChild(remove_el);
        }
        var remove_li = document.getElementById("report_links");

        if (remove_li != null) {
            head.removeChild(remove_li);
        }

        if (level1 != '') { jQuery("a#"+level1).removeClass('current'); }
        level1 = controller;
        jQuery("a#"+level1).addClass('current');

        var report_links = document.createElement("div");
        report_links.setAttribute("id", "report_links");
        head.appendChild(report_links);
        //var ul_link = document.createElement("ul");

        for (var i = 0; i < links_array.length; i++) {
            //var li_link = document.createElement("li");
            if (ind == Indexes[0]) {
                if ((i!=1)&& (i != 3) && (i != 5) && (i!=10) && (i!=12) && (i!=13)) continue;
            }
            else if (ind == Indexes[1]) {
                if ((i==2) || (i==3) || (i == 4)) continue;
            }
            var a_link = document.createElement("a");
            var a_text = document.createTextNode(links_array[i]);
            a_link.appendChild(a_text);
            a_link.setAttribute("id", "rep_" + i);
            a_link.setAttribute("class", "button_bar");
            report_links.appendChild(a_link);
            //li_link.appendChild(a_link);
            //ul_link.appendChild(li_link);
            jQuery("a#rep_" + i).click(function (event) {
                var remove_el = document.getElementById("cur_reports");
                if (remove_el != null) {
                    report_body.removeChild(remove_el);
                }
                //event.target.setAttribute("class", "current");
                var reports_cur = document.createElement("div");
                reports_cur.setAttribute("id", "cur_reports");
                var id_arr = event.target.id.split('_');
                var header = document.createElement("h1");
                var header_text = document.createTextNode(links_array[id_arr[1]] + ((ind == Indexes[0]) ? ", K$" : " "));
                header.appendChild(header_text);

                var download_link = document.createElement("a");
                download_link.setAttribute("href", "javascript:populateIframe('frame1','"+filenames[id_arr[1]]+"')");
                var dwnld_img = document.createElement("img");
                dwnld_img.setAttribute("src","/wp-content/uploads/icons/32/file_extension_xls.png");
                download_link.appendChild(dwnld_img);
                var download_text = document.createTextNode("Download in Excel");
                download_link.appendChild(download_text);
                download_link.setAttribute("style", "position: relative; float: right; top: -12px");
                header.appendChild(download_link);
                reports_cur.appendChild(header);
                report_body.appendChild(reports_cur);

                if (level2 != '') { jQuery("a#"+level2).removeClass('current'); }
                level2 = event.target.id;
                jQuery("a#"+level2).addClass('current');

                jQuery.ajax({
                    type: "GET",
                    url: url,
                    data: "report_id=" + (parseInt(id_arr[1])+ind),
                    statusCode: {
                        404: function () {
                            alert("page not found");
                        },
                        500: function () {
                            alert("internal server error");
                        }
                    },
                    success: function (result) {
                        if (result.type == 'error') {
                            alert('error');
                            return(false);
                        }
                        else {
                            //do nothing
                            //alert('success function');
                            var rep_tab = document.createElement('div');
                            rep_tab.setAttribute("id", "multi_sheet");
                            rep_tab.innerHTML = result;
                            reports_cur.appendChild(rep_tab);
                            var sheet_links = document.getElementById("sheet_links");
                            if (sheet_links != null) {
                                //it means we have multi-sheet xls, process 'title_%x%' as controllers, sheet'%x%' as views
                                var j = 0;
                                var sheet_title = document.getElementById("title_" + j);
                                while (sheet_title != null) {
                                    jQuery("table#sheet" + j).hide();
                                    jQuery("a#title_" + j).click(function (ev) {
                                        var tar = ev.target.id.split('_');
                                        jQuery("table#sheet" + tar[1]).show();
                                        if (level3 != '') {
                                            jQuery("a#"+level3).removeClass('current');
                                            var tar1 = level3.split('_');
                                            jQuery("table#sheet" + tar1[1]).hide();
                                        }
                                        level3 = ev.target.id;
                                        jQuery("a#"+level3).addClass('current');
                                    });
                                    j++;
                                    sheet_title = document.getElementById("title_" + j);
                                }
                                //jQuery("col").css({'width': 'auto'});
                                //$('td.column0').css({'text-align':'left'});
                                var table_all = $('table'),
                                    //tWidth = 980,
                                    cols = table_all.find('col'),
                                    firstColumnWidth = 200;
                                    //colWidth = (tWidth - firstColumnWidth) / cols.size();
                                //table_all.width(tWidth);
                                //table_all.css({'position': 'relative', 'left': '-20px'});
                                cols.width(60);
                                //$('td').width(colWidth);
                                $('td').css({'width':60, 'text-align': 'center', 'vertical-align': 'middle'});
                                $('col.col0').width(firstColumnWidth);
                                $('td.column0').css({'width':firstColumnWidth, 'text-align':'left'});
                            }
                            else {
                                var table = $('table#sheet0'),
                                    tWidth = 980,
                                    cols = table.find('col'),
                                    firstColumnWidth = 250,
                                    colWidth = (tWidth - firstColumnWidth) / cols.size();
                                table.width(tWidth);
                                table.css({'position': 'relative', 'left': '-20px'});
                                cols.width(colWidth);
                                //$('td').width(colWidth);
                                $('td').css({'width':colWidth, 'text-align': 'center', 'vertical-align': 'middle'});
                                if (ind ==Indexes[2]) {
                                    table.width(firstColumnWidth + 60*4);
                                    cols.width(60);
                                    $('td').css({'width':60, 'text-align': 'center', 'vertical-align': 'middle'});
                                }
                                $('col.col0').width(firstColumnWidth);
                                $('td.column0').css({'width':firstColumnWidth, 'text-align':'left', 'padding-left': '5px'});

                            }
                            if (ind == Indexes[0]) {
                                $('td').each(function (i) {
                                    //var temp = /\d+,\d+/.test(this.childNodes[0].data);
                                    if (this.childNodes[0] != undefined) {
                                        if ((/^\d+/.test(this.childNodes[0].data) || /^-\d+/.test(this.childNodes[0].data)) && !/%/.test(this.childNodes[0].data)) {
                                            //(!/^[A-Za-z](?=.)/.test(this.childNodes[0].data)) && ((/\d+,\d+/.test(this.childNodes[0].data)) || (!/\d+(?=%)/.test(this.childNodes[0].data)))
                                            var kVal = Math.round(parseFloat(this.childNodes[0].data.replace(/,/g, '')) / 1000);
                                            this.childNodes[0].data = kVal.format();
                                        } else if (/^x(?=\d+)/.test(this.childNodes[0].data)) {
                                            var jVal = /(^x)(\d+)/.exec(this.childNodes[0].data);
                                            this.childNodes[0].data = jVal[2];
                                        } else if (/^x(?!.)/.test(this.childNodes[0].data)) {
                                            this.childNodes[0].data = " ";
                                        }
                                    }
                                })
                            }
                        }
                    }
                })
            })
        }
        //report_links.appendChild(ul_link);
    });
}
