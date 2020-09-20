<?php
    require("php/database.php");
    if (isset($_GET)){
        $name = $_GET['name'];
        $message = $_GET['message'];
        $date = date('Y-m-d H:i:s');
        
        $connection = ConnectToDatabase();
        $sql = "INSERT INTO `MESSAGE` (`name`, `message`, `date`) ";
        $sql .= "VALUES ('" . $name . "', '" . $message . "', '" . $date . "');";
        
        if (!(mysqli_query($connection, $sql))){
            echo("Error" . mysqli_error($connection));
        }
        else {
            echo("Message sent");
        }
    }
?>
