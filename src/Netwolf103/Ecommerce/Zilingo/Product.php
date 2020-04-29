<?php
namespace Netwolf103\Ecommerce\Zilingo;

/**
 * Product class.
 *
 * @author Zhang Zhao <netwolf103@gmail.com>
 */
class Product
{
	/**
	 * Product data
	 *
	 * @var array
	 */
	private $data = [];

	/**
	 * PHPExcel object
	 *
	 * @var \PHPExcel
	 */
	protected $excel;

	/**
	 * Read csv file
	 *
	 * @param string $csvFile
	 */
	function __construct(string $csvFile)
	{
		$csv = array_map('str_getcsv', file($csvFile));

		$data = [];
	    array_walk($csv, function(&$item) use (&$data) {
			$data[$item[0]] = $item[1];
	    });

	    $this->data 	= $data;
	    $this->excel 	= new \PHPExcel();
	}

	/**
	 * Magic method "getXX"
	 *
	 * @param  string $name
	 * @param  array $arguments
	 * @return mixed
	 */
    public function __call($name, $args) 
    {
    	$value = '';

        if (substr($name, 0, 3) == 'get') {
        	$field = ucwords(substr($name, 3));
        	$value = $this->data[$field] ?? '';

            if (!$value) {
                $config = $this->loadConfig();
                $value = $config[$field] ?? '';
            }            
        }

        return $value;
    }

    /**
     * Return image urls
     *
     * @return array
     */
    public function getImageUrls(): array
    {
    	$urls = [];

	    array_walk($this->data, function($item, $key) use (&$urls) {
	    	if (strstr($key, 'Image')) {
	    		$urls[] = $item;
	    	}
	    });

	    return $urls; 	
    }

    /**
     * Return special price
     *
     * @param  float $multiple
     * @param  float $suffix
     * @param  float $minPrice
     * @return float
     */
    public function getSpecialPrice(float $multiple = 5.0, float $suffix = 0.95, float $minPrice = 100.0): float
    {
    	$price = $this->data['Price'] ?? '';
    	$price = explode(' ', $price);
    	$price = $price[1] ?? 0;
    	$price = $price * $multiple;

    	if ($price <= 0) {
    		$price = $minPrice;
    	}

    	return floor($price) + $suffix;
    }

    /**
     * Return price
     *
     * @param  float  $multiple
     * @param  float  $suffix
     * @return float
     */
    public function getPrice(float $multiple = 2.2, float $suffix = 0.95): float
    {
    	$specialPrice = $this->getSpecialPrice();

    	return floor($specialPrice * $multiple) + $suffix;
    }

    /**
     * Return stock
     *
     * @param  int|integer $stock
     * @return int
     */
    public function getStock(int $stock = 10): int
    {
    	return $stock;
    }

    /**
     * Return shipping fee
     *
     * @param  float  $price
     * @return float
     */
    public function getShipFee(float $price = 20.0): float
    {
    	return $price;
    }

    /**
     * Return sizes
     *
     * @param  float  $min
     * @param  float  $max
     * @param  float  $step
     * @return array
     */
    public function getSizes(float $min = 4.0, float $max = 12.0, float $step = 0.5): array
    {
        return range($min, $max, $step);
    }

    /**
     * Return material
     *
     * @return string
     */
    public function getMaterial(): string
    {
        return '金属';
    }

    /**
     * Create a excel file
     *
     * @param  string $filename
     * @return self
     */
    public function saveExcel(string $filename)
    {
    	$docTitle = sprintf('Zilingo Product - %s', $this->getSku());

		$this->excel
			->getProperties()
			->setTitle($docTitle)
			->setSubject($docTitle)
			->setDescription($docTitle)
			->setKeywords($docTitle)
			->setCategory($docTitle);

        $this->excel->getActiveSheet()->setTitle('WJEART-F,ARR-S,US');

        //Row 1
        $this->excel
            ->getActiveSheet()
            ->setCellValue('A1', $docTitle)
            ->mergeCells('A1:U1')
            ->getStyle("A1:BC1")->applyFromArray([
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                ],
                'font' => [
                    'bold' => true,
                    'size' => 12,
                    'name' => 'Calibri (Body)'
                ]
            ])
        ;

        //Row 2
        $this->excel
            ->getActiveSheet()
            ->setCellValue('D2', 'Version:12')
            ->setCellValue('E2', 'Locale: zh-Hans')
            ->setCellValue('F2', 'CHV: 1.3.3')
            ->setCellValue('G2', 'PV:B2C')
            ->setCellValue('H2', 'AT:REGULAR')
            ->setCellValue('I2', 'FTS:F,ARR-S,US')
            ->getStyle('D2:I2')->applyFromArray([
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                ],
                'font' => [
                    'bold' => true,
                    'size' => 12,
                    'name' => 'Calibri (Body)'
                ]
            ])
        ;

        //Row 3
        $this->excel
            ->getActiveSheet()
            ->setCellValue('P3', 'Enter the Stock')
            ->mergeCells('P3:AL3')
            ->setCellValue('AN3', 'Weight （选填）')
            ->mergeCells('AN3:AO3')
            ->setCellValue('AP3', 'Dimensions （选填）')
            ->mergeCells('AP3:AS3')
            ->setCellValue('AT3', 'Bulk Image Upload (Provide at least one link)')
            ->mergeCells('AT3:BC3')                      
            ->getStyle('P3:BC3')->applyFromArray([
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                ],
                'font' => [
                    'bold' => true,
                    'size' => 12,
                    'name' => 'Calibri (Body)'
                ]
            ])
        ;        

	    array_walk($this->getHeader(), function($item, $cell) {
			$this->excel
				->getActiveSheet()
				->setCellValue($cell.'4', $item)
			;
	    });
        $this->excel
            ->getActiveSheet()
            ->mergeCells('F4:H4')
            ->mergeCells('J4:L4')
            ->getStyle('A4:BC4')->applyFromArray([
                'alignment' => [
                    'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                ],
                'font' => [
                    'bold' => true,
                    'size' => 12,
                    'name' => 'Calibri (Body)'
                ]
            ])
        ;      

        $this->excel
            ->getActiveSheet()
            ->setCellValue('A5', $this->getColor())
            ->setCellValue('B5', $this->getName())
            ->setCellValue('C5', $this->getDesc())
            ->setCellValue('D5', '')
            ->setCellValue('E5', $this->getSku())
            ->setCellValue('F5', $this->getGemType())
            ->setCellValue('I5', $this->getPolishing())
            ->setCellValue('J5', $this->getMaterial())
            ->setCellValue('M5', $this->getStoneShape())
            ->setCellValue('N5', $this->getMetalWeight())
            ->setCellValue('O5', $this->getWarehouse())
            ->setCellValue('P5', '')
            ->setCellValue('Q5', '')
            ->setCellValue('R5', 10)
            ->setCellValue('S5', 10)
            ->setCellValue('T5', 10)
            ->setCellValue('U5', 10)
            ->setCellValue('V5', 10)
            ->setCellValue('W5', 10)
            ->setCellValue('X5', 10)
            ->setCellValue('Y5', 10)
            ->setCellValue('Z5', 10)
            ->setCellValue('AA5', 10)
            ->setCellValue('AB5', 10)
            ->setCellValue('AC5', 10)
            ->setCellValue('AD5', 10)
            ->setCellValue('AE5', 10)
            ->setCellValue('AF5', 10)
            ->setCellValue('AG5', 10)
            ->setCellValue('AH5', 10)
            ->setCellValue('AI5', '')
            ->setCellValue('AJ5', '')
            ->setCellValue('AK5', '')
            ->setCellValue('AL5', '')
            ->setCellValue('AM5', $this->getPrice())
            ->setCellValue('AN5', $this->getWeightUnit())
            ->setCellValue('AO5', $this->getWeightValue())
            ->setCellValue('AP5', '')
            ->setCellValue('AQ5', '')
            ->setCellValue('AR5', '')
            ->setCellValue('AS5', '')
        ;

        $imageUrls = $this->getImageUrls();
        foreach(['AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC'] as $i => $cell) {
            $imageUrl = $imageUrls[$i] ?? '';
            $this->excel
                ->getActiveSheet()
                ->setCellValue($cell.'5', $imageUrl)
            ;
        }

        $objWriter = \PHPExcel_IOFactory::createWriter($this->excel, 'Excel2007');
        $objWriter->save($filename);        

		return $this;
    }

    /**
     * Return excel header
     *
     * @return array
     */
    protected function getHeader(): array
    {
    	$header = [
    		'A' => '颜色',
    		'B' => '名称',
    		'C' => '描述',
    		'D' => 'Brand （选填）',
    		'E' => '卖家SKU编号',
    		'F' => '原生宝石类型 （选填）',
    		'G' => 'NULL',
    		'H' => 'NULL',
    		'I' => '抛光 （选填）',
    		'J' => '材质 (Select at least one)',
    		'K' => 'NULL',
    		'L' => 'NULL',
    		'M' => '宝石形状 （选填）',
    		'N' => '金属重量 （选填）',
            'O' => 'Warehouse',
            'P' => '3 美国',
            'Q' => '3.5美国',
            'R' => '4 美国',
            'S' => '4.5美国',
            'T' => '5 美国',
            'U' => '5.5美国',
            'V' => '6 美国',
            'W' => '6.5美国',
            'X' => '7 美国',
            'Y' => '7.5美国',
            'Z' => '8 美国',
            'AA' => '8.5美国',
            'AB' => '9 美国',
            'AC' => '9.5美国',
            'AD' => '10美国',
            'AE' => '10.5美国',
            'AF' => '11美国',
            'AG' => '11.5美国',
            'AH' => '12美国',
            'AI' => '12.5美国',
            'AJ' => '13美国',
            'AK' => '13.5美国',
            'AL' => '14美国',
            'AM' => '价格 (in CNY)',
            'AN' => 'Weight Unit',
            'AO' => 'Value',
            'AP' => 'Dimension Unit',
            'AQ' => 'Length',
            'AR' => 'Breadth',
            'AS' => 'Height',
            'AT' => '图片 1',
            'AU' => '图片 2',
            'AV' => '图片 3',
            'AW' => '图片 4',
            'AX' => '图片 5',
            'AY' => '图片 6',
            'AZ' => '图片 7',
            'BA' => '图片 8',
            'BB' => '图片 9',
            'BC' => '图片 10',
    	]; 	

    	return $header;
    }

    /**
     * Return config
     *
     * @return array
     */
    protected function loadConfig(): array
    {
    	$config = dirname(dirname(dirname(dirname(dirname(__FILE__))))) . '/config.csv';

    	$data = [];

    	if (file_exists($config)) {
			$csv = array_map('str_getcsv', file($config));

			$data = [];
		    array_walk($csv, function(&$item) use (&$data) {
				$data[$item[0]] = $item[1];
		    });
    	}

    	return $data;
    }    
}