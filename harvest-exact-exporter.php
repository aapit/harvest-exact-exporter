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


	protected $_dateColumnLabels = array('Date');

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

		$this->_formatDateColumns();
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
		$newSrcContent = array_map(array($this, '_stripNumber'), $oldSrcContent);
		$newDestContent = array_map(array($this, '_extractNumber'), $oldSrcContent);
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

	protected function _stripNumber($value) {
		$pattern = '/(\D+)( ?\(\d+\))/i';
		return trim(preg_replace($pattern, '$1', $value));
	}

	protected function _extractNumber($value) {
		if (strpos($value, '(') === false) {
			return null;
		}
		$pattern = '/[a-z \/\&]*(\s*\((\d+)\))?/i';
		return preg_replace($pattern, '$2', $value);
	}

	protected function _formatDateColumns() {
		$columns = array();

		foreach ($this->_dateColumnLabels as $label) {
			$columns[] = $this->_getHeaderColumn($label);
		}

		foreach ($columns as $column) {
			$this->_formatDateColumn($column);
		}

		//print_r($this->_getHeaderRow());
		//exit();


	}

	/**
 	 * @param String $column The column letter
 	 */
	protected function _formatDateColumn($column) {
		$this->_getSheet()
    		->getStyle(
				$this->_getColumnContentCoordinates($column)
			)
    		->getNumberFormat()
    		->setFormatCode(PHPExcel_Style_NumberFormat::FORMAT_DATE_DDMMYYYY)
		;

		// FORMAT_DATE_DMYSLASH
		// FORMAT_DATE_DMYMINUS
		// http://www.cmsws.com/examples/applications/phpexcel/Documentation/API/PHPExcel_Style/PHPExcel_Style_NumberFormat.html#constFORMAT_DATE_DMYSLASH
	}

	/**
 	 * @param String $column	The column letter
 	 * @return String			The coordinates of the content part of the column,
 	 *							sans header. For instance: "A2:A599"
 	 */
	protected function _getColumnContentCoordinates($column) {
		return 
			$column . self::FIRST_CONTENT_ROW
			. ':'
			. $column . $this->_getSheet()->getHighestRow()
		;
	}

	/**
 	 * Retrieves the header column that corresponds with given label.
 	 * @return String The column letter
 	 */
	protected function _getHeaderColumn($label) {
		$headerRow = $this->_getHeaderRow();

		return array_search($label, $headerRow);
	}

	/**
 	 * Retrieves the cell values in the header row.
 	 * @return Array array(
 	 *					'A' => 'Column Name 1',
 	 *					'B' => 'Column Name 2'
 	 *				 )
 	 */
	protected function _getHeaderRow() {
		$headerCells = array();

		$row = $this->_getSheet()->getRowIterator(self::HEADER_ROW)->current();
		$cellIterator = $row->getCellIterator();
		$cellIterator->setIterateOnlyExistingCells(false);

		foreach ($cellIterator as $cell) {
    		$headerCells[$cell->getColumn()] = $cell->getValue();
		}

		return $headerCells;
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
