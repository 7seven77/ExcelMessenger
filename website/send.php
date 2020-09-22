<?php
    require("php/database.php");
    if (isset($_GET)){
        $sender = $_GET['sender'];
        $recipient = $_GET['recipient'];
        $message = $_GET['message'];
        $date = date('Y-m-d H:i:s');
        
        $connection = ConnectToDatabase();
        $sql = "INSERT INTO `MESSAGE` (`sender`, `recipient`, `message`, `date`) ";
        $sql .= "VALUES ('" . $sender . "', '" . $recipient . "', '" . $message . "', '" . $date . "');";
        
        if (!(mysqli_query($connection, $sql))){
            echo("Error" . mysqli_error($c));
        }
        else {
            echo("Message sent");
        }
    }
?>
