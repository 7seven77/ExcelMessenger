<?php
    require("php/database.php");
    if (isset($_GET)){
        $sender = $_GET['sender'];
        $recipient = $_GET['recipient'];
        $message = $_GET['message'];
        $date = date('Y-m-d H:i:s');
        
        $connection = ConnectToDatabase();
        $sql = "CALL `addMessage`(\"" . $sender . "\", \"" . $recipient . "\", \"" . $message . "\")";
        
        if (!(mysqli_query($connection, $sql))){
            echo("Error" . mysqli_error($connection));
        }
        else {
            echo("Message sent");
        }
    }
?>
