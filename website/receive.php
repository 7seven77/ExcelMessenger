<?php
    require("php/database.php");
    if (isset($_GET)){
        $sender = $_GET['sender'];
        $recipient = $_GET['recipient'];
        
        $connection = ConnectToDatabase();
        $sql = "SELECT * FROM `MESSAGE` WHERE (";
        $sql .= "`sender` = '" . $sender . "' and ";
        $sql .= "`recipient` = '" . $recipient . "') ";
        $sql .= "OR (";
        $sql .= "`sender` = '" . $recipient . "' and ";
        $sql .= "`recipient` = '" . $sender . "') ";
        $sql .= "ORDER BY `date` DESC;";
        $result = mysqli_query($connection, $sql);
        
        if (mysqli_num_rows($result) > 0) {
            while($row = mysqli_fetch_assoc($result)) {
                echo(json_encode($row));
            }
        }
        else {
          echo("No results");
        }
    }
?>
