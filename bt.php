<?php



require_once('excel/PHPExcel.php');
require_once('excel/PHPExcel/IOFactory.php');
$db = new SQLite3('budget');
$cols = array( 'account'=>1, 'entity'=>1, 'dept'=>1,'subdept'=>1,'natural'=>1,'desc'=>1,'subnatural'=>1,
	'act_2ya'=>2,
	'act_1ya'=>3,
	'plan_1ya'=>4,
	'plan' =>5 );

initDb();

function initDb() {
	global $db;
	$cmd = <<<EOD
		drop table targets;
		create table targets(id integer PRIMARY KEY,
		  account text,
		  entity text,
		  dept text,
		  subdept text,
		  natural text,
		  subnatural text,
		  desc text,
		  act_2ya real,
		  act_1ya real,
		  plan_1ya real,
		  plan real real);
EOD;
	if ( !	$db->exec( $cmd)) {
		die($error);
	}

}
$xls = PHPExcel_IOFactory::load("budget1.xlsx");
$sht = $xls->getActiveSheet();
$rows = $sht->getHighestRow();
$colnm = $sht->getHighestColumn();
$maxcols = PHPExcel_Cell::columnIndexFromString( $colnm );

for($i=0;$i<$rows;$i++) {
  $cell = $sht->getCellByColumnAndRow(0,$i)->getValue();
	if ( strpos($cell, "100-") !== FALSE ) {
		processRow( $i );
	}
}

function processRow( $index ) {
	global $db;

	$sql = <<<SQL
	insert into targets(account,entity,dept,subdept,natural,subnatural,desc,act_2ya,act_1ya,plan_1ya,plan)
	values (
SQL;
	$sql .= "'" . getVal($index,'account') . "',"; 
	$sql .= "'" . getVal($index,'entity') . "',"; 
	$sql .= "'" . getVal($index,'dept') . "',"; 
	$sql .= "'" . getVal($index,'subdept') . "',"; 
	$sql .= "'" . getVal($index,'natural') . "',"; 
	$sql .= "'" . getVal($index,'subnatural') . "',"; 
	$sql .= "'" . $db->escapeString(getVal($index,'desc')) . "',"; 
	$sql .= "'" . getVal($index,'act_2ya') . "',"; 
	$sql .= "'" . getVal($index,'act_1ya') . "',"; 
	$sql .= "'" . getVal($index,'plan_1ya') . "',"; 
	$sql .= "'" . getVal($index,'plan') . "'"; 
	$sql .= ");";
	// echo($sql);
	$db->exec($sql);
	echo( getVal($index,'account') . "\n");
}

function getVal( $index, $name ) {
	global $sht, $cols;
	$col = $cols[$name];
	$cell = $sht->getCellByColumnAndRow( $col-1, $index )->getValue() ;
	if ($col==1) {
		$parts = explode( "-", substr($cell,0,17));
		switch($name) {
			case 'account':
				return substr($cell,0,17);
			case 'desc':
				return substr($cell,19);
			case 'entity':
				return $parts[0];
				break;
			case 'dept':
				return substr($parts[1],0,2);
				break;
			case 'subdept':
				return substr($parts[1],2,2);
				break;
			case 'natural':
				return $parts[2];
				break;
			case 'subnatural':
				return $parts[3];
				break;
		}
	}	
	return $cell;
}
