<?php
/**
 * Created by PhpStorm.
 * User: buphmin
 * Date: 12/6/15
 * Time: 4:26 PM
 */

require_once(realpath(__DIR__ . '/..') . "/vendor/autoload.php");
require_once(realpath(__DIR__ . '/..') . '/src/Container.php');
require_once(realpath(__DIR__ . '/..') . '/src/Transfer.php');

$container = new \PHPExcelDataFiles\Container(realpath(__DIR__ . '/..').'/example/files/data_inventory.txt', 'id');

$toContainer = new \PHPExcelDataFiles\Container(realpath(__DIR__ . '/..') . '/example/files/test.xlsx', 'id');


$sync = new \PHPExcelDataFiles\Transfer();
$sync->setFromContainer($container);
$sync->setToContainer($toContainer);
$sync->syncSheets(0, 'main');

$sync->saveContainers();


