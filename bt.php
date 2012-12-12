<?php



require_once('excel/PHPExcel.php');
require_once('excel/PHPExcel/IOFactory.php');
$db = new SQLite3('budget');
$cols = array( 'account'=>1, 'entity'=>1, 'dept'=>1,'subdept'=>1,'natural'=>1,'desc'=>2,'subnatural'=>1,
	'act_2ya'=>3,
	'act_1ya'=>4,
	'plan_1ya'=>5,
	'plan' =>6 );

initDb();

function initDb() {
	global $db;
	$cmd = <<<EOD
		drop table targets;
		create table targets(id integer PRIMARY KEY,
		  account text,
		  entity integer,
		  dept integer,
		  subdept integer,
		  natural integer,
		  subnatural integer,
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
	echo( getVal($index,'entity'));
	echo( getVal($index,'dept'));
	echo( getVal($index,'subdept'));
	echo( getVal($index,'natural'));
	echo( getVal($index,'subnatural'));
	echo("\n");
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
			case 'entity':
				return $parts[0];
				break;
			case 'dept':
				return substr($parts[1],0,2);
				break;
			case 'subdept':
				return substr($parts[1],3,2);
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
