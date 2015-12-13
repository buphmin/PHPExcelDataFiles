<?php
/**
 * Created by PhpStorm.
 * User: buphmin
 * Date: 12/6/15
 * Time: 4:05 PM
 */

namespace PHPExcelDataFiles;


/**
 * This Class uses the Container object to easily sync sheets
 * from each workbook or delimited file based on a primary key.
 * Class Transfer
 * @package PHPExcelDataFiles
 */
class Transfer
{
    /**
     * @var Container
     */
    private $toContainer;

    /**
     * @var Container
     */
    private $fromContainer;

    /**
     * @param \PHPExcelDataFiles\Container $container
     */
    public function setToContainer(Container $container)
    {
        $this->toContainer = $container;
    }

    /**
     * @param \PHPExcelDataFiles\Container $container
     */
    public function setFromContainer(Container $container)
    {
        $this->fromContainer = $container;
    }

    /**
     * @param $filePath
     * @param $primaryKey
     */
    public function setToContainerFromPath($filePath, $primaryKey)
    {
        $this->toContainer = new Container($filePath, $primaryKey);
    }

    /**
     * @param $filePath
     * @param $primaryKey
     */
    public function setFromContainerFromPath($filePath, $primaryKey)
    {
        $this->fromContainer = new Container($filePath, $primaryKey);
    }

    /**
     * Syncs the data from a sheet in the fromContainer to a sheet
     * in the toContainer.
     * @param $fromSheetName
     * @param $toSheetName
     * @throws \PHPExcel_Exception
     */
    public function syncSheets($fromSheetName, $toSheetName)
    {
        $fromHeader = $this->fromContainer->getHeaderBySheet($fromSheetName);
        $toHeader = $this->toContainer->getHeaderBySheet($toSheetName);

        $fromKeyList = $this->fromContainer->getPrimaryKeyList($fromSheetName);
        $toKeyList = $this->toContainer->getPrimaryKeyList($toSheetName);

        // Uses the header and primary key dictionaries to
        // map between each sheet.
        foreach($fromHeader as $header => $column)
        {
            foreach($fromKeyList as $id => $row)
            {
                // don't try and transfer if the key doesn't exist in the destination
                if(isset($toHeader[$header]) and isset($toKeyList[$id]))
                {
                    $fromCellValue = $this
                        ->fromContainer
                        ->getSheet($fromSheetName)
                        ->getCellByColumnAndRow($column, $row)
                        ->getValue();


                    $this
                        ->toContainer
                        ->getSheet($toSheetName)
                        ->getCellByColumnAndRow($toHeader[$header], $toKeyList[$id])
                        ->setValue($fromCellValue);
                }
            }
        }
    }

    /**
     * Tries to save each container
     * @throws \Exception
     */
    public function saveContainers()
    {
        $this->toContainer->saveBook();
        $this->fromContainer->saveBook();
    }
}