<?php

	// ****** START CONFIG ******

	// относительный путь к excel-файлу
	$file_xls = 'ost.xls';

	// относительный путь к файлу с датой последней обработки
	$file_last = 'last-check.txt';

	// доступ к БД
	$db_host  = 'localhost';
	$db_user  = '';
	$db_pass  = '';
	$db_name  = '';

	// номер категории
	define('CAT_ID', 314);

	// номера колонок
	define('COL_UID', 1);
	define('COL_SKU', 2);
	define('COL_CATEGORY', 3);
	define('COL_NAME', 4);
	define('COL_FULL_NAME', 5);
	define('COL_DESCRIPTION', 6);
	define('COL_CONTENT', 7);
	define('COL_PRICE', 8);
	define('COL_UNIT', 9);
	define('COL_STOCK', 10);
	define('COL_IMAGE', 11);

	// выводить отладочную инфу на экран true (да) / false (нет)
	$debug = true;

	// ****** END CONFIG ********


	// инициализация
	ini_set('display_errors', 'on');
	error_reporting(E_ALL);
	mb_internal_encoding('utf-8');
	$dir = dirname(__FILE__) . '/';
	require_once($dir . 'excel/reader.php');
	require_once($dir . 'excel/oleread.php');
	$file_xls = $dir . $file_xls;
	$file_last = $dir . $file_last;

	$db = new mysqli($db_host, $db_user, $db_pass, $db_name);
	if ($db->connect_error) {
		die('Connect Error (' . $db->connect_errno . ') ' . $db->connect_error);
	}

	// проверяем дату последнего обновления
	clearstatcache();
	$mtime = filemtime($file_xls);
	if (is_file($file_last)) {
		if ($mtime <= file_get_contents($file_last)) {
			if ($debug) echo 'No changes';
			exit;
		}
	}
	
	// обнуляем у всех stock
	$db->query('update `s_variants` set `stock`=0') or die($db->error);

	// построчно читаем excel
	$data = new Spreadsheet_Excel_Reader();
	$data->setOutputEncoding('cp1251');
	$data->read($file_xls);
	$rows_count = $data->sheets[0]['numRows'];
	for ($row = 2; $row <= $rows_count; $row++) {
		$cols = array();
		for($col = 1; $col <= 11; $col++) {
			$value = isset($data->sheets[0]['cells'][$row][$col])
				? $data->sheets[0]['cells'][$row][$col]
				: '';
			$value = iconv('cp1251', 'utf-8', $value);
			if ($col == COL_SKU && preg_match('/^\d+$/', $value)) {
				$value = str_pad($value, 5, '0', STR_PAD_LEFT);
			}
			$value = $db->real_escape_string($value);
			$cols[$col] = $value;
		}

		if ($debug) echo "\nSKU: " . $cols[COL_SKU] . "\n";
		
		// есть такой товар?
		$result = $db->query(
			'select id from `s_variants` where ' .
			'`sku`="' . $cols[COL_SKU] . '"'
		) or die($db->error);
		$item = $result->fetch_assoc();
		if ($item) {
			// есть - апдейтим price и stock
			$query = 
				'update `s_variants` set ' .
				'`price`="' . $cols[COL_PRICE] . '", ' .
				'`stock`="' . $cols[COL_STOCK] . '" ' .
				'where `id`=' . $item['id'];
			$db->query($query) or die($db->error);
			if ($debug) echo "UPDATE: $query\n";
		} else {
			// нет - добавляем новый товар

			// s_products
			$arr = array(
				'url'              => translit($cols[COL_FULL_NAME]),
				'name'             => $cols[COL_FULL_NAME],
				'annotation'       => '',
				'body'             => '',
				'visible'          => 1,
				'meta_title'       => $cols[COL_FULL_NAME],
				'meta_keywords'    => $cols[COL_FULL_NAME],
				'meta_description' => $cols[COL_FULL_NAME],
				'external_id'      => '',
			);
			db_insert('s_products', $arr);
			$prod_id = $db->insert_id;

			// s_variants
			$arr = array(
				'product_id'    => $prod_id,
				'sku'           => $cols[COL_SKU],
				'name'          => '',
				'price'         => $cols[COL_PRICE],
				'stock'         => $cols[COL_STOCK],
				'position'      => 0,
				'attachment'    => '',
				'external_id'   => '',
			);
			db_insert('s_variants', $arr);
			$var_id = $db->insert_id;

			// s_products_categories
			$arr = array(
				'product_id'  => $prod_id,
				'category_id' => CAT_ID,
				'position'    => 0,
			);
			db_insert('s_products_categories', $arr);
		}
	}

	// сохраняем дату последнего обновления
	file_put_contents($file_last, $mtime);

	function db_insert($table, $arr) {
		global $db, $debug;

		$chunks = array();
		foreach($arr as $k=>$v) {
			$chunks[] = "`$k`='$v'";
		}
		$query = "insert into `$table` set " . implode(',', $chunks);
		$db->query($query) or die($db->error);
		if ($debug) echo "INSERT: $query\n";
	}

	function translit($s) {
		$s = preg_replace('/[^\w\-]/u', '-', $s);
		$rus = array(
			'а','б','в','г','д','е','ё','ж','з','и','й','к','л','м','н','о','п','р','с','т','у','ф','х','ц','ч','ш','щ','ъ','ы','ь','э','ю','я',
			'А','Б','В','Г','Д','Е','Ё','Ж','З','И','Й','К','Л','М','Н','О','П','Р','С','Т','У','Ф','Х','Ц','Ч','Ш','Щ','Ъ','Ы','Ь','Э','Ю','Я',
		);
		$lat = array(
			'a','b','v','g','d','e','yo','zh','z','i','j','k','l','m','n','o','p','r','s','t','u','f','h','c','ch','sh','shch','','y','','e','yu','ya',
			'A','B','V','G','D','E','Yo','Zh','Z','I','J','K','L','M','N','O','P','R','S','T','U','F','H','C','Ch','Sh','Shch','','Y','','E','Yu','Ya',
		);
		$s = str_replace($rus, $lat, $s);
		$s = preg_replace('/\-{2,}/', '-', $s);
		return(trim($s, '-'));
	}
