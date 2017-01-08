(function ($) {
	$.excelHelpers = (function () {
		var obj = {};
		obj.defaults = {
			cols: 5,
			rows: 10
		};
		obj.getCellBounds = function ($cell) {
			var offset = $cell.offset();
			return {
				left: offset.left,
				top: offset.top,
				right: offset.left + $cell.outerWidth(),
				bottom: offset.top + $cell.outerHeight(),
				width: $cell.outerWidth(),
				height: $cell.outerHeight(),
			}
		};
		obj.setSelection = function ($instance, $cell) {
			//var bounds = [$cell.offset().left, $cell.offset().top, $cell.outerWidth(), $cell.outerHeight()];
			var bounds = obj.getCellBounds($cell);
			var $selection = $instance.find('.selection');
			$selection.find('.left').css({ left: bounds.left - 1, top: bounds.top - 1, width: 2, height: bounds.height + 2 })
			$selection.find('.top').css({ left: bounds.left - 1, top: bounds.top - 1, width: bounds.width + 2, height: 2 })
			$selection.find('.bottom').css({ left: bounds.left - 1, top: bounds.bottom, width: bounds.width + 2, height: 2 })
			$selection.find('.right').css({ left: bounds.right, top: bounds.top - 1, width: 2, height: bounds.height + 2 })
		};
		obj.setEditable = function ($instance, $cell) {
			//var bounds = obj.getCellBounds($cell);
			$cell.addClass('active').find('input.editbox').focus().select();
			// var $editbox = $instance.find('.editbox');
			// $editbox.css({ left: bounds.left + 1, top: bounds.top + 1, width: bounds.width - 1, height: bounds.height - 1 })
			// 	.show().val($cell.text()).focus().select();
		};
		obj.setFormulaBar = function ($instance, $cell) {
			var $formularInput = $instance.find('input.formula');
			var $cellInput = $cell.find('input.editbox');
			$formularInput.val($cellInput.attr('data-formula'));
		};
		obj.getLetterAt = function (n, lower) {
			return String.fromCharCode((lower ? 97 : 65) + n - 1);
		};
		obj.getAddressAt = function (i, j) {
			return obj.getLetterAt(j) + String(i);
		};
		obj.getCellByAddress = function ($instance, address) {
			// var splitted = address.split(/(\d+)/).filter(Boolean);
			// var rowIndex = parseInt(splitted[1]);
			// var colIndex = splitted[0].charCodeAt(0); 
			return $instance.find('[data-cell-address="' + address + '"]')
		};
		return obj;
	})()
	
	$.fn.excelSheet = function (options) {
		options = $.extend({}, $.excelHelpers.defaults, options);
		return this.each(function () {
			var $instance = $(this);
			
			// formulabar
			var $formulabarDiv = $('<div></div>').appendTo($instance);
			var $formulaInput = $('<input type="text" class="formula" />')
				.on('keydown', function (evt) {
					evt = evt ? evt : window.event;
					var charCode = (evt.which) ? evt.which : evt.keyCode;
					if (charCode == 13) {
						var $cell = $instance.data('activeCellObject');
						var $cellInput = $cell.find('input.editbox');
						$cellInput.attr('data-formula', $(this).val());
					}
				})
				.appendTo($formulabarDiv)
			
			// create table
			var $table = $('<table>').appendTo($instance);
			var $thead = $('<thead>').appendTo($table);
			var $tbody = $('<tbody>').appendTo($table);
			for (var i = 0; i <= options.rows; i++) {
				if (i == 0) {
					var $tr = $('<tr>').appendTo($thead);
					
					for (var j = 0; j <= options.cols; j++) {
						if (j == 0)
							$('<th>').appendTo($tr)
						else
							$('<th>').text($.excelHelpers.getLetterAt(j)).appendTo($tr)
					}
				} else {
					var $tr = $('<tr>').appendTo($tbody);
					for (var j = 0; j <= options.cols; j++) {
						if (j == 0) {
							$('<th>').text(i).appendTo($tr);
						} else {
							var $td = $('<td>')
								.attr('data-cell-address', $.excelHelpers.getAddressAt(i, j))
								.attr('data-cell-rowIndex', i)
								.attr('data-cell-colIndex', j)
								.click(function (e) {
									$instance.data('activeCellObject', $(this));
									$.excelHelpers.setSelection($instance, $(this));
									$.excelHelpers.setFormulaBar($instance, $(this));
									//console.debug($.excelHelpers.getCellByAddress($instance, $(this).attr('data-cell-address')))
									//console.debug([$(this).data('cellAddress'), $(this).data('cellRowIndex'), $(this).data('cellColIndex')])
								}).dblclick(function (event) {
									$.excelHelpers.setEditable($instance, $(this));
									event.stopPropagation();
									return false
								})
								.appendTo($tr);
								
							var $input = $('<input id="' + $.excelHelpers.getAddressAt(i, j) + '" type="text" class="editbox" />')
								.blur(function (e) {
									var $this = $(this);
									$this.parent()
										.removeClass('active')
										.find('span').text($this.val());
								})
								.on('keydown', function (evt) {
									evt = evt ? evt : window.event;
									var charCode = (evt.which) ? evt.which : evt.keyCode;
									if (charCode == 13) $(this).trigger('blur')
								})
								.appendTo($td);
								
							var $span = $('<span></span>').appendTo($td);
						}
					}
				}
			}
			
			// selection
			var $selection = $('<div class="selection">')
				.append($('<div>').addClass('top'))
				.append($('<div>').addClass('right'))
				.append($('<div>').addClass('bottom'))
				.append($('<div>').addClass('left'))
				.appendTo($instance);
				
			// editbox
			// var $editbox = $('<input type="text" class="editbox">')
			// 	.blur(function (e) {
			// 		// var $cell = $instance.data('activeCellObject');
			// 		// if ($cell) {
			// 		// 	$cell.text($editbox.val())
			// 		// }
			// 		// $editbox.val('').hide().css({ left: 0, top: 0 })
			// 	})
			// 	.on('keydown', function (evt) {
			// 		evt = evt ? evt : window.event;
      		// 		var charCode = (evt.which) ? evt.which : evt.keyCode;
			// 		if (charCode == 13) $(this).trigger('blur')
			// 	})
			// 	.appendTo($instance)
			
		})
	}
})(jQuery)