<?php

class Report {
    var $user;
    var $password;
    var $dbname;
    var $server_url;
    var $user_id;
    var $relation_names;
    var $relation_keys;
    var $search_keys;
    var $data_arr;

    function constructor($uname, $upass, $dbname, $serv_url, $rel_names, $rel_keys, $sch_keys){
        $this->user = $uname;
        $this->password = $upass;
        $this->dbname = $dbname;
        $this->server_url = $serv_url;
        $this->relation_names = $rel_names;
        $this->relation_keys = $rel_keys;
        $this->search_keys = $sch_keys;
    }

    function connect_xmlrpc() {

       if(isset($_COOKIE["user_id"]) == true)  {
           if($_COOKIE["user_id"]>0) {
           //echo $_COOKIE["user_id"];
           }
       }

       $sock = new xmlrpc_client($this->server_url.'common');

       $msg = new xmlrpcmsg('login');
       $msg->addParam(new xmlrpcval($this->dbname, "string"));
       $msg->addParam(new xmlrpcval($this->user, "string"));
       $msg->addParam(new xmlrpcval($this->password, "string"));
       $resp =  $sock->send($msg);

       $val = $resp->value();
       $this->user_id = $val->scalarval();

       setcookie("user_id",$this->user_id,time()+3600);
       if($this->user_id > 0) {
           return $this->user_id;
       }else{
           return -1;
       }
    }

    function get_xml_data($rep_id) {

        $sock = new xmlrpc_client($this->server_url.'object');
        //$client->return_type = 'phpvals';// or 'xml';
        $sock->return_type = 'xml';

        //echo "User_id = ".$this->user_id.", Relation = ".$this->relation_names[$rep_id];
        $msg = new xmlrpcmsg('execute');
        $msg->addParam(new xmlrpcval($this->dbname, "string"));
        $msg->addParam(new xmlrpcval($this->user_id, "int"));
        $msg->addParam(new xmlrpcval($this->password, "string"));
        $msg->addParam(new xmlrpcval($this->relation_names[$rep_id], "string"));
        $msg->addParam(new xmlrpcval("search", "string"));
        $msg->addParam(new xmlrpcval($this->search_keys[$rep_id], "array"));
        $resp =  $sock->send($msg);

        if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();
        //return $resp->faultCode();
        $val = $resp->value();
        //var_dump($val);
        $v = xmlrpc_decode($val);
        //return $user_id;
        //echo "Responce = ".var_dump($v);
        $ids_read = array();
        foreach($v AS $index => $value) {
            $ids_read[]=new xmlrpcval($value, "int");
        }

        //echo $this->relation_names[$rep_id];
        //var_dump($this->relation_keys[$rep_id]);
        $msg_read = new xmlrpcmsg('execute');
        $msg_read->addParam(new xmlrpcval($this->dbname, "string"));
        $msg_read->addParam(new xmlrpcval($this->user_id, "int"));
        $msg_read->addParam(new xmlrpcval($this->password, "string"));
        $msg_read->addParam(new xmlrpcval($this->relation_names[$rep_id], "string"));
        $msg_read->addParam(new xmlrpcval("read", "string"));
        $msg_read->addParam(new xmlrpcval($ids_read, "array"));
        //$msg_read->addParam(new xmlrpcval($this->relation_keys[$rep_id], "array"));

        $resp =  $sock->send($msg_read);
        //echo "Whatup!";
        if ($resp->faultCode())
            echo "Got error on ->send(read)".$resp->faultCode();
        $str_val = $resp->value();
        $val = xmlrpc_decode($str_val);
        //var_dump($val);
        $this->data_arr[] = $val;
    //return json_encode($val);
    } //standard query function end

    //custom query function (for VB and HR reports)
    function get_xml_data_c($rep_id, $ids_read) {
        //
        $sock = new xmlrpc_client($this->server_url.'object');
        //$client->return_type = 'phpvals';// or 'xml';
        $sock->return_type = 'xml';
        //echo $this->relation_names[$rep_id];
        //var_dump($this->relation_keys[$rep_id]);
        $msg_read = new xmlrpcmsg('execute');
        $msg_read->addParam(new xmlrpcval($this->dbname, "string"));
        $msg_read->addParam(new xmlrpcval($this->user_id, "int"));
        $msg_read->addParam(new xmlrpcval($this->password, "string"));
        $msg_read->addParam(new xmlrpcval($this->relation_names[$rep_id], "string"));
        $msg_read->addParam(new xmlrpcval("read", "string"));
        $msg_read->addParam(new xmlrpcval($ids_read, "array"));
        //$msg_read->addParam(new xmlrpcval($this->relation_keys[$rep_id], "array"));

        $resp =  $sock->send($msg_read);
        //echo "Whatup!";
        if ($resp->faultCode())
            echo "Got error on ->send(read)".$resp->faultCode();
        $str_val = $resp->value();
        $val = xmlrpc_decode($str_val);
        //var_dump($val);
        $this->data_arr[] = $val;
    }

    //get_xml_data_reg - for all Regular Revenue Reports
    function get_xml_data_reg($rep_id) {

        $sock = new xmlrpc_client($this->server_url.'object');
        //$client->return_type = 'phpvals';// or 'xml';
        $sock->return_type = 'xml';

        $msg = new xmlrpcmsg('execute');
        $msg->addParam(new xmlrpcval($this->dbname, "string"));
        $msg->addParam(new xmlrpcval($this->user_id, "int"));
        $msg->addParam(new xmlrpcval($this->password, "string"));
        $msg->addParam(new xmlrpcval($this->relation_names[$rep_id], "string"));
        $msg->addParam(new xmlrpcval("search", "string"));
        $msg->addParam(new xmlrpcval($this->search_keys[0], "array"));
        $resp =  $sock->send($msg);

        if ($resp->faultCode())
            echo "Got error on ->send()".$resp->faultCode();
        //return $resp->faultCode();
        $val = $resp->value();
        //var_dump($val);
        $v = xmlrpc_decode($val);
        //return $user_id;
        $ids_read = array();
        foreach($v AS $index => $value) {
            $ids_read[]=new xmlrpcval($value, "int");
        }

        //echo $this->relation_names[$rep_id];
        //var_dump($this->relation_keys[$rep_id]);
        $msg_read = new xmlrpcmsg('execute');
        $msg_read->addParam(new xmlrpcval($this->dbname, "string"));
        $msg_read->addParam(new xmlrpcval($this->user_id, "int"));
        $msg_read->addParam(new xmlrpcval($this->password, "string"));
        $msg_read->addParam(new xmlrpcval($this->relation_names[$rep_id], "string"));
        $msg_read->addParam(new xmlrpcval("read", "string"));
        $msg_read->addParam(new xmlrpcval($ids_read, "array"));
        //$msg_read->addParam(new xmlrpcval($this->relation_keys[0], "array"));

        $resp =  $sock->send($msg_read);
        //echo "Whatup!";
        if ($resp->faultCode())
            echo "Got error on ->send(read)".$resp->faultCode();
        $str_val = $resp->value();
        $val = xmlrpc_decode($str_val);
        //var_dump($val);
        $this->data_arr[] = $val;
    //return json_encode($val);
    } //get_xml_data_reg - for all Regular Revenue Reports

} //end of class description

?>