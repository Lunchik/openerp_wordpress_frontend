<?php
/*
Template Name: DXB ERP Reports template
*/
?>
<?php get_header(); ?>

<style>

#head {
    overflow: auto;
}

#reports_head {
    width: 220px;
    float: left;
}

#reports_head ul {
    list-style: none;
    margin: 0;
    padding: 0;
}

#reports_head li {
    /*padding-bottom: 2px;*/
}

#reports_head a {
    border-top-left-radius: 0px;
    border-top-right-radius: 0px;
    border-bottom-right-radius: 0px;
    border-bottom-left-radius: 0px;
}

#reports_head li a {
    height: 32px;
    voice-family: '"}"';
    voice-family: inherit;
    height: 24px;
    text-decoration: none;
}

#reports_head li a:link, #reports_head li a:visited {
    color: #036B7E;
    display: block;
    background: transparent;
    padding: 8px 0 0 10px;
}

#reports_head li a:hover, #reports_head li a.current {
    color: #036B7E;
    background:  #DCEFF2;
    padding: 8px 0 0 10px;
}
#report_links {
    float: left;
    width: 700px;
    border-left: 1px dotted #ccc;
    background: transparent;
    font-size: 16px;
    color: #000;
    border-bottom: 1px dotted #ccc;
}

#report_links a {
    padding: 5px;
    margin: 0 5px 0 5px;
    width: auto;
    display: inline-block;
    font-size: 14px;
    color: #036B7E;
}

#report_links a.button_active {
    color: #036B7E;
    background: #DCEFF2;
}


#report_links a:hover, #report_links a.current {
    color: #036B7E;
    background:  #DCEFF2;
}


#sheet_links a {
    padding: 5px;
    margin: 0 5px 0 5px;
    width: auto;
    display: inline-block;
    font-size: 12px;
}

#sheet_links a:hover, #sheet_links a.current {
    color: #036B7E;
    background:  #DCEFF2;
}

#cur_reports h1 {
    color: #036B7E;
}

</style>

<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script src="http://tgtoil.com/wp-content/themes/MyCuisine/xmlrpc_lib/event_handler.js" type="text/javascript"></script>

<?php get_template_part('includes/breadcrumbs'); ?>

    <div class="container fullwidth">
        <div id="content" class="clearfix">
            <div id="left-area">

    <div id="head">
        <div name="reports" id="reports_head">
            <ul>
            <li><a href="#" class="rep_button" id="rev">Revenue Reports</a></li>
            <li><a href="#" class="rep_button" id="ast">Asset Reports</a></li>
            <li><a href="#" class="rep_button" id="hrr">HR Reports</a></li>
            </ul>
        </div>
        <div name="reports_links" id="report_links" >
        </div>
    </div>

    </br>
    <div name="report_body" id="reports_body" >
	<iframe id="frame1" style="display:none"></iframe>
    </div>


 </div>
             <!-- end #left-area -->

         </div>
         <!-- end #content -->
         <div id="bottom-shadow"></div>
     </div> <!-- end .container -->

 <?php get_footer(); ?>