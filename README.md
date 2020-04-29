# netwolf103/ecommerce-zilingo
Create product excel file of upload Zilingo

# Code maintainers
![头像](https://avatars3.githubusercontent.com/u/1772352?s=100&v=4)
------------
Zhang Zhao <netwolf103@gmail.com>
------------
Wechat: netwolf103

## Require
	"php": "^7.0",
	"phpoffice/phpexcel": "1.8.2"

## Install
composer require netwolf103/ecommerce-zilingo

## Usage
```PHP
<?php
require_once "vendor/autoload.php";

use Netwolf103\Ecommerce\Zilingo\Product;

$varDir = sprintf('%s/var/1688/%s', dirname(dirname(__FILE__)), date('Y-m-d'));
if (!is_dir($varDir)) {
	printf("%s dir not exists.", $varDir);
	exit;
}

if ($handle = opendir($varDir)) {
    while (false !== ($file = readdir($handle))) {
		if ($file == '.' || $file == '..') {
			continue;
		}

		if(substr($file, -3) != 'csv') {
			continue;
		}

		$file 		= $varDir .'/'. $file;
		$fileExcel 	= $file .'.zilingo.xlsx';

		if (file_exists($fileExcel)) {
			continue;
		}

        $product = new Product($file);
        $product->saveExcel($fileExcel);
    }

    closedir($handle);
}
```