<?php
    function ConnectToDatabase() {
        $server = "";
        $user = "";
        $pass = "";
        $name = "";
        $connection = mysqli_connect($server, $user, $pass, $name);
        if (mysqli_connect_errno()){
            return NULL;
        }
        return $connection;
    }
    
    function GetRow($connection, $query){
        $result = mysqli_query($connection, $query);
        if (mysqli_num_rows($result) == 0){
            return NULL;
        }
        return $row = mysqli_fetch_assoc($result);
    }
?>
