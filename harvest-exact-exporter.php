<?php

require 'PHPExcel/Classes/PHPExcel.php';
date_default_timezone_set('Europe/Amsterdam');

$path = $argv[1];
$sheet = new HarvestSheet($path);
$sheet->splitColumn('B', 'Client Code');
$sheet->output();

class HarvestSheet {
	const ERROR_NO_FILE =
		'Please provide a filename for the source, as first parameter.';
	const HEADER_ROW = 1;
	const FIRST_CONTENT_ROW = 2;

	/**
 	 * @var PHPExcel
 	 */
	protected $_excelDoc;

	protected $_path;

	/**
 	 * Store the first blank column.
 	 */
	protected $_destColumn;


	public function __construct($path) {
		$this->_path = $path;
		$this->_excelDoc = $this->_openFile();
		$this->_destColumn = chr(ord($this->_getSheet()->getHighestColumn()) +1);
		//$this->_parseContent();
	}

	public function output() {
		$objWriter = new PHPExcel_Writer_Excel5($this->_excelDoc);
		$objWriter->save(
			dirname($this->_path) . DIRECTORY_SEPARATOR
			. basename($this->_path, '.xlsx')
			. '.xls'
		);
	}

	public function splitColumn($srcColumn, $destName) {
		$this->_setDestinationColumnName($destName);

		$oldSrcContent = $this->_getColumnContent($srcColumn);
		$clean = function($value) {
			$pattern = '/(\D+)( ?\(\d+\))/i';
			return trim(preg_replace($pattern, '$1', $value));
		};
		$extractNumber = function($value) {
			if (strpos($value, '(') === false) {
				return null;
			}
			$pattern = '/[a-z \/\&]*(\s*\((\d+)\))?/i';
			return preg_replace($pattern, '$2', $value);
		};

		$newSrcContent = array_map($clean, $oldSrcContent);
		$newDestContent = array_map($extractNumber, $oldSrcContent);
		$newSrcColumn = $this->_makeColumn($newSrcContent);
		$newDestColumn = $this->_makeColumn($newDestContent);

		$this->_getSheet()->fromArray(
			$newSrcColumn,
			null,
			$srcColumn . self::FIRST_CONTENT_ROW	
		);

		$this->_getSheet()->fromArray(
			$newDestColumn,
			null,
			$this->_destColumn . self::FIRST_CONTENT_ROW
		);
	}

	protected function _getColumnContent($srcColumn) {
		$highestRow = $this->_getSheet()->getHighestRow();
		$from = $srcColumn . self::FIRST_CONTENT_ROW;
		$to = $srcColumn . $highestRow;

		$columnData = $this->_getSheet()
			->rangeToArray(
				$from . ':' . $to
			)
		;

		$flatten = function(&$value) {
			$value = $value[0];
		};
		array_walk($columnData, $flatten);

		return $columnData;
	}

	protected function _makeColumn(array $values) {
		return array_chunk($values, 1);
	}

	protected function _setDestinationColumnName($destName) {
		$this->_getSheet()->setCellValue($this->_destColumn . self::HEADER_ROW, $destName);
	}

	/**
 	 * @return PHPExcel
 	 */
	protected function _openFile() {
		if (empty($this->_path)) {
			throw new Exception(self::ERROR_NO_FILE);
		}

		$excelDoc = PHPExcel_IOFactory::load($this->_path);
		return $excelDoc;
	}

	protected function _getSheet() {
		return $this->_excelDoc->getActiveSheet();
	}
}
