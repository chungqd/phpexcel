<?php
function connection() {
	try {
		$dbh = new PDO('mysql:host=localhost;dbname=phpexcel;charset=utf8', 'root', 'Chung@1994');
		return $dbh;

	} catch (PDOException $e) {
		print"Error!: ".$e->getMessage()."<br/>";
		die();
	}
}

function disconnection($conn) {
	$conn = null;
}

?>
