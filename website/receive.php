<?php
    require("database.php");
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
        
        $separator = "¬&£@*^%";
        
        if (mysqli_num_rows($result) > 0) {
            while($row = mysqli_fetch_assoc($result)) {
                echo($row['sender'] . $separator . $row['recipient'] . $separator . $row['message'] . $separator . $row['date'] . $separator . $separator);
            }
        }
        else {
          echo("No results");
        }
    }
?>

