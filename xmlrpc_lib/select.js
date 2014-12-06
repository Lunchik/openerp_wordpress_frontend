/**
 * Created by nelia on 5/18/14.
 */


$ = jQuery.noConflict();

var reportsArray;
var url = 'http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/select_action.php';

var report_ids = [1,2,3,4,5,6,7,8,9,10/*,11,12,13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26*/];

var report_names = [
    "Revenue by Client/Cost Centre", //1
    "Revenue by Cost Centre/Country", //2
    "Revenue by Service/Client", //3
    "Revenue by Service/Country", //4
    "Revenue by Service/Category 1", //5
    "Revenue by Service/Category 2", //6
    "Revenue by Client/Country", //7
    "Revenue by Service/Operator", //8
    "Revenue by Operator/Country", //9
    /*"Revenue per Job/Country", //10
    "Revenue per Job/Client", //11
    "Revenue per Job/Operator", //12
    "Revenue per Job/Category", *///13
    "Revenue vs. Target YTD"/*, //10
    "Revenue vs. Target Jan 2014", //15
    "Revenue vs. Target Feb 2014", //16
    "Revenue vs. Target Mar 2014", //17
    "Revenue vs. Target Apr 2014", //18
    "Revenue vs. Target May 2014", //19
    "Revenue vs. Target Jun 2014", //20
    "Revenue vs. Target Jul 2014", //21
    "Revenue vs. Target Aug 2014", //22
    "Revenue vs. Target Sep 2014", //23
    "Revenue vs. Target Oct 2014", //24
    "Revenue vs. Target Nov 2014", //25
    "Revenue vs. Target Dec 2014" *///26
];

//country cache
var Country_Cache = new Array();

/*
DOCUMENT LOADED EVENT
 */
jQuery(document).ready(function() {

    $('div#loading-image').show();

    /*
    1. create an "array" of calls
    2. data for the array of ajax calls is an array [1,2,3,4,5,6,7,8,9,10,11,12,13] (some of them are deprecated)
     */

    var calls = [];
    var current = 0;

    var reportsArray = [];
    var url = 'http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/select_action.php';

    var report_ids = [1,2,3,4,5,6,7,8,9,10];

    var report_names = [
        "Sale by Client/Cost Centre", //1
        "Sale by Cost Centre/Country", //2
        "Sale by Service/Client", //3
        "Sale by Service/Country", //4
        "Sale by Service/Category 1", //5
        "Sale by Service/Category 2", //6
        "Sale by Client/Country", //7 -- create one more object that saves the revenue by country
        "Sale by Service/Operator", //8
        "Sale by Operator/Country"/*, //9
        "Revenue per Job/Country", //10
        "Revenue per Job/Client", //11
        "Revenue per Job/Operator", //12
        "Revenue per Job/Category" *///13

    ];

    var target_reps = [
        "Revenue vs. Target YTD", //14
        "Revenue vs. Target Jan 2014", //15
        "Revenue vs. Target Feb 2014", //16
        "Revenue vs. Target Mar 2014", //17
        "Revenue vs. Target Apr 2014", //18
        "Revenue vs. Target May 2014", //19
        "Revenue vs. Target Jun 2014", //20
        "Revenue vs. Target Jul 2014", //21
        "Revenue vs. Target Aug 2014", //22
        "Revenue vs. Target Sep 2014", //23
        "Revenue vs. Target Oct 2014", //24
        "Revenue vs. Target Nov 2014", //25
        "Revenue vs. Target Dec 2014" //26
    ];

    /***************************************
    for (var i=0; i<report_ids.length; i++) {
        calls.push(jQuery.ajax({
            type: "GET",
            url: url,
            data: "report_id="+report_ids[i],
            statusCode: {
                404: function() {
                    alert( "page not found" );
                },
                500: function() {
                    alert("internal server error");
                }
            },
            dataType: "json"
        }))
    }
    ****************************************/

    var select_handler = document.getElementById("report_id");

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
                    //enabling the <option>
                    if (current==0) {
                        var inopt = document.getElementById("init0");
                        if (inopt != null) {
                            inopt.remove();
                        }
                    }
                    if (current <= 9) {
                        var opt = document.createElement("option");
                        var optText = document.createTextNode(report_names[current]);
                        opt.appendChild(optText);
                        opt.setAttribute("value", "" + (current + 1) + "");
                        select_handler.appendChild(opt);
                        if (current == 6) fill_up_the_Country_Cache(serverResponse);
                    }
                    else {
                        for (var i=0; i<target_reps_id.length; i++) {
                            var opt = document.createElement("option");
                            var optText = document.createTextNode(target_reps[i]);
                            opt.appendChild(optText);
                            opt.setAttribute("value", "" + (target_reps_id[i]) + "");
                            select_handler.appendChild(opt);
                        }
                    }

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


    jQuery("select#report_id").change(function() {
        var remove_el = document.getElementById("report_container");
        if (remove_el != null) {
            remove_el.remove();
        }
        //alert ("Report change event triggered!");
        var report_id = $(this).val();
        //alert (report_id);
        if (report_id != 0) {
            //alert (report_id);
            //$('div#loading-image').show();
            //show the report (draw it using the previously queried data
            if (report_id <= 13) {
                create_table_Sales_by_Month(reportsArray[report_id - 1], report_id);
            }
            else if (report_id <= 26) {
                create_table_Revenue_by_Target(reportsArray[13], report_id);
            }
        }
        else alert ("Please, select valid report from the list.");
    });
});


/*
*
* MTM and YTY Target reports
*
 */

function create_table_Revenue_by_Target(linesArray, report_id) {

    //report_id = 14 - YTD
    //            15 - Jan , etc.

    var mon_arr = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "des"];
    var month_no = report_id-15;
    //prepare the data - sort the array
    linesArray.sort(compare3);

    var element = document.getElementById("reports_div");

    var report_container = document.createElement("div");
    report_container.setAttribute("id", "report_container");

    var header = document.createElement("h1");
    var header_text;
    var rep_name = report_names[report_id-1];

    header_text = document.createTextNode(rep_name+"\n");
    header.appendChild(header_text);
    report_container.appendChild(header);
    //start creating a table
    var tformat = ["Country", "Actual", "Target", "Variance"];
    var table = document.createElement("table");
    table.setAttribute("id", "sales_month");

    fill_theader(table, tformat);

    var tbody = document.createElement("tbody");
    var trow;

    if (report_id == 14) {
        //YTD
        var country = '';
        var cou_val = 0;
        for (var i=0; i<linesArray.length; i++) {
            //
            if (country == '') {
                country = linesArray[i].country_id[1];
                cou_val = linesArray[i].tar_mo;
            }
            else if (strcmp(linesArray[i].country_id[1], country)){
                //same country
                cou_val += linesArray[i].tar_mo;
            }
            else {
                //new country
                trow = document.createElement("tr");
                var cache = Country_Cache.filter(function(o){
                    return strcmp(o.defcol_one_id, country);
                });
                fill_target_ytd_tbrow(trow, country, cache[0].months[13], cou_val);
                tbody.appendChild(trow);
                //new country init
                country = linesArray[i].country_id[1];
                cou_val = linesArray[i].tar_mo;
            }
        }
        trow = document.createElement("tr");
        var cache = Country_Cache.filter(function(o){
            return strcmp(o.defcol_one_id, country);
        });
        fill_target_ytd_tbrow(trow, country, cache[0].months[12], cou_val);
        tbody.appendChild(trow);
    }
    else { //months
        for (var i=0; i<linesArray.length; i++) {
            //choose the lines that have the [report_id] month and display them as a row + add a Country_Cache
            if (mon_arr[month_no] == linesArray[i].mon) {
                //fill the table. else - nothing
                trow = document.createElement("tr");
                var value = linesArray[i].country_id[1];
                var cache = Country_Cache.filter(function(o){
                    return strcmp(o.defcol_one_id, value);
                });
                fill_target_tbrow(trow, linesArray[i], cache, month_no);
                tbody.appendChild(trow);
            }
        }
    }
    table.appendChild(tbody);
    report_container.appendChild(table);
    element.appendChild(report_container);
    return 0;
}

function strcmp ( str1, str2 ) {

    var size1 = str1.charCodeAt ( 0 );
    var size2 = str2.charCodeAt ( 0 );

    return  ( size1 == size2 ) ? 1 : 0;
}

function compare3(a,b) {
    if (a.country_id[1] < b.country_id[1])
        return -1;
    if (a.country_id[1] > b.country_id[1])
        return 1;
    return 0;
}

function remove_quart(Cache) {
    //
    Cache.months.splice(3, 1);
    Cache.months.splice(6,1);
    Cache.months.splice(9, 1);
    return Cache;
}

function fill_target_ytd_tbrow(trow, country, act, target){
    var td;
    var td_text;

    td = document.createElement("td");
    td.setAttribute("class", "title");
    td_text = document.createTextNode(country);
    td.appendChild(td_text);
    trow.appendChild(td);

    td = document.createElement("td");
    td_text = document.createTextNode(act);
    td.appendChild(td_text);
    trow.appendChild(td);

    td = document.createElement("td");
    td_text = document.createTextNode(target);
    td.appendChild(td_text);
    trow.appendChild(td);

    var diff = target - act;
    td = document.createElement("td");
    td_text = document.createTextNode(diff.toString());
    td.appendChild(td_text);
    trow.appendChild(td);

}

function fill_target_tbrow(trow, line, Cache, month_no){
    var td;
    var td_text;

    td = document.createElement("td");
    td.setAttribute("class", "title");
    td_text = document.createTextNode(line.country_id[1]);
    td.appendChild(td_text);
    trow.appendChild(td);

    td = document.createElement("td");
    td_text = document.createTextNode(Cache[0].months[month_no]);
    td.appendChild(td_text);
    trow.appendChild(td);

    td = document.createElement("td");
    td_text = document.createTextNode(line.tar_mo);
    td.appendChild(td_text);
    trow.appendChild(td);

    var diff = parseInt(line.tar_mo, 10) - parseInt(Cache[0].months[month_no],10);
    td = document.createElement("td");
    td_text = document.createTextNode(diff.toString());
    td.appendChild(td_text);
    trow.appendChild(td);

}

function fill_up_the_Country_Cache(linesArray) {

        //sorting
        linesArray.sort(compare1);
        linesArray.sort(compare2);

        //create the l_total, g_total
        var l_total = new ReportLine();
        var g_total = new ReportLine();
        g_total.init_gtotal();
        var cc = '';
        for (var i = 0; i<linesArray.length; i++) {
            //create MY objects using linesArray[i]
            var line = new ReportLine();
            line.fill(linesArray[i]);
            if (cc=='') {
                //cc=line.defcol_two_id; l_total = init; display the client
                cc = line.defcol_two_id;
                l_total.init_total(line);
            }
            else if (cc == line.defcol_two_id) {
                //l_total - sum with line; display the client
                l_total.sum(line);
            }
            else {
                //g_total sum with l_total; display the cc with ltotal; cc=line.defcol_two_id;
                //display the client; init total with line
                var temp = l_total;
                l_total = new ReportLine();
                Country_Cache.push(remove_quart(temp));
                g_total.sum(l_total);
                cc=line.defcol_two_id;
                l_total.init_total(line);
            }
        }
    var temp = l_total;
    delete l_total;
    Country_Cache.push(remove_quart(temp));
    return 0;
}

/*
**
* These are the classes and functions for the standard monthly reports
**
 */
function create_table_Sales_by_Month(linesArray, report_id){

    var element = document.getElementById("reports_div");

    var report_container = document.createElement("div");
    report_container.setAttribute("id", "report_container");

    var header = document.createElement("h1");
    var header_text;
    var rep_name = report_names[report_id-1];
    //if report_id = 1, then Sales by Client/Cost Centre - need to change defcol_one and defcol_two
    if (report_id == 1) {
        //print header Revenue by Client/Cost Centre +++ switch field values
        //now switching values
        for (var i = 0; i < linesArray.length; i++) {
            var val_container = linesArray[i].defcol_one_id;
            linesArray[i].defcol_one_id = linesArray[i].defcol_two_id;
            linesArray[i].defcol_two_id = val_container;
        }
    }

    header_text = document.createTextNode(rep_name+"\n");
    header.appendChild(header_text);
    report_container.appendChild(header);
    //start creating a table
    var tformat = [rep_name, "Jan", "Feb", "Mar", "Q1", "Apr", "May", "Jun", "Q2", "Jul", "Aug", "Sep", "Q3", "Oct", "Nov", "Dec", "Q4", "YTD"];
    var table = document.createElement("table");
    table.setAttribute("id", "sales_month");

    fill_theader(table, tformat);

    var tbody = document.createElement("tbody");
    var trow;

    //sorting
    linesArray.sort(compare1);
    linesArray.sort(compare2);

    //create the l_total, g_total
    var l_total = new ReportLine();
    var g_total = new ReportLine();
    g_total.init_gtotal();
    var cc = '';
    for (var i = 0; i<linesArray.length; i++) {
        //create MY objects using linesArray[i]
        trow = document.createElement("tr");
        var line = new ReportLine();
        line.fill(linesArray[i]);
        if (cc=='') {
            //cc=line.defcol_two_id; l_total = init; display the client
            cc = line.defcol_two_id;
            l_total.init_total(line);
            fill_tbrow(trow, line);
        }
        else if (cc == line.defcol_two_id) {
            //l_total - sum with line; display the client
            l_total.sum(line);
            fill_tbrow(trow, line);
        }
        else {
            //g_total sum with l_total; display the cc with ltotal; cc=line.defcol_two_id;
            //display the client; init total with line
            g_total.sum(l_total);
            fill_tbrow_total(trow, l_total);
            tbody.appendChild(trow);
            cc=line.defcol_two_id;
            trow = document.createElement("tr");
            fill_tbrow(trow, line);
            l_total.init_total(line);
        }
        tbody.appendChild(trow);
    }
    trow = document.createElement("tr");
    g_total.sum(l_total);
    fill_tbrow_total(trow, l_total);
    tbody.appendChild(trow);
    trow = document.createElement("tr");
    fill_tbrow_total(trow, g_total);
    tbody.appendChild(trow);
    table.appendChild(tbody);
    report_container.appendChild(table);
    element.appendChild(report_container);
    return 0;
}

function ReportLine() {
    this.defcol_one_id ='';
    this.defcol_two_id ='';
    this.months = new Array(18).join('0').split('').map(parseFloat);

    this.fill = function(line){
        this.defcol_one_id =line.defcol_one_id[1];
        this.defcol_two_id =line.defcol_two_id[1];
        this.months[0]=line.jan;
        this.months[1]=line.feb;
        this.months[2]=line.mar;
        this.months[3]=line.q1;
        this.months[4]=line.apr;
        this.months[5]=line.may;
        this.months[6]=line.jun;
        this.months[7]=line.q2;
        this.months[8]=line.jul;
        this.months[9]=line.aug;
        this.months[10]=line.sep;
        this.months[11]=line.q3;
        this.months[12]=line.oct;
        this.months[13]=line.nov;
        this.months[14]=line.des;
        this.months[15]=line.q4;
        this.months[16]=line.total;

    };

    this.init_gtotal = function () {
        this.defcol_one_id = "TOTAL";
        //months - already array of 0
    };

    this.init_total = function (line) {
        this.defcol_one_id =line.defcol_two_id;
        this.months = line.months;
    };

    this.sum = function(line){
        //this.defcol_one_id ='';
        //this.defcol_two_id ='';

        for (var i=0; i<this.months.length; i++) {
            this.months[i] += line.months[i];
        }
        /*this.months[0] +=line.jan;
        this.months[1] +=line.feb;
        this.months[2] +=line.mar;
        this.months[3] +=line.q1;
        this.months[4] +=line.apr;
        this.months[5] +=line.may;
        this.months[6] +=line.jun;
        this.months[7] +=line.q2;
        this.months[8] +=line.jul;
        this.months[9] +=line.aug;
        this.months[10] +=line.sep;
        this.months[11] +=line.q3;
        this.months[12] +=line.oct;
        this.months[13] +=line.nov;
        this.months[14] +=line.des;
        this.months[15] +=line.q4;
        this.months[16] +=line.total; */
    };

}


//fill tables
function fill_theader(tab_obj, format) {
    var thead = document.createElement("thead");
    var trow = document.createElement("tr");
    trow.setAttribute("class", "header");
    var td;
    var td_text;

    for (var i=0; i<format.length; i++) {
        td = document.createElement("th");
        if (i==0) td.setAttribute("class", "title");
        td_text = document.createTextNode(format[i]);
        td.appendChild(td_text);
        trow.appendChild(td);
    }
    thead.appendChild(trow);
    tab_obj.appendChild(thead);
}

function fill_tbrow(trow, line){
    var td;
    var td_text;

    td = document.createElement("td");
    td.setAttribute("class", "title");
    td_text = document.createTextNode(line.defcol_one_id);
    td.appendChild(td_text);
    trow.appendChild(td);

    for (var i=0; i<line.months.length; i++) {
        td = document.createElement("td");
        if ((i==3) || (i==7) || (i==11) || (i==15) || (i==16)) {
            td.setAttribute("class", "q_total");
        } else {
            td.setAttribute("class", "num");
        }
        td_text = document.createTextNode((line.months[i]).formatMoney(0, '.', ','));
        td.appendChild(td_text);
        trow.appendChild(td);
    }
}

function fill_tbrow_total(trow, line){
    var td;
    var td_text;

    trow.setAttribute("class", "total");
    td = document.createElement("td");
    td.setAttribute("class", "title");
    td_text = document.createTextNode(line.defcol_one_id);
    td.appendChild(td_text);
    trow.appendChild(td);

    for (var i=0; i<line.months.length; i++) {
        td = document.createElement("td");
        td.setAttribute("class", "num");
        td_text = document.createTextNode((line.months[i]).formatMoney(0, '.', ','));
        td.appendChild(td_text);
        trow.appendChild(td);
    }
}

//these are the functions for sorting the lines by 2 different columns
function compare1(a,b) {
    if (a.defcol_one_id[1] < b.defcol_one_id[1])
        return -1;
    if (a.defcol_one_id[1] > b.defcol_one_id[1])
        return 1;
    return 0;
}
function compare2(a,b) {
    if (a.defcol_two_id[1] < b.defcol_two_id[1])
        return -1;
    if (a.defcol_two_id[1] > b.defcol_two_id[1])
        return 1;
    return 0;
}


Number.prototype.formatMoney = function(c, d, t){
    var n = this,
        c = isNaN(c = Math.abs(c)) ? 2 : c,
        d = d == undefined ? "." : d,
        t = t == undefined ? "," : t,
        s = n < 0 ? "-" : "",
        i = parseInt(n = Math.abs(+n || 0).toFixed(c)) + "",
        j = (j = i.length) > 3 ? j % 3 : 0;
    return s + (j ? i.substr(0, j) + t : "") + i.substr(j).replace(/(\d{3})(?=\d)/g, "$1" + t) + (c ? d + Math.abs(n - i).toFixed(c).slice(2) : "");
};

/*var xml_arr=result_arr[1].split("<?xml version='1.0'?>");
 if (window.DOMParser)
 {
 parser=new DOMParser();
 xmlDoc=parser.parseFromString(xml_arr[1],"text/xml");
 console.log(xmlDoc);
 }
 else // Internet Explorer
 {
 xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
 xmlDoc.async=false;
 xmlDoc.loadXML(xml_arr[1]);
 }*/