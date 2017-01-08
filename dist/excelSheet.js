/**************************
 * excelSheet 0.1
 * 2017, January
 * Mostafa Rowghanian
 * ************************/
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
			var bounds = obj.getCellBounds($cell);
			var $selection = $instance.find('.selection');
			$selection.find('.left').css({ left: bounds.left - 1, top: bounds.top - 1, width: 2, height: bounds.height + 2 })
			$selection.find('.top').css({ left: bounds.left - 1, top: bounds.top - 1, width: bounds.width + 2, height: 2 })
			$selection.find('.bottom').css({ left: bounds.left - 1, top: bounds.bottom, width: bounds.width + 2, height: 2 })
			$selection.find('.right').css({ left: bounds.right, top: bounds.top - 1, width: 2, height: bounds.height + 2 })
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
		if (typeof(options) == 'string') {
			var $instance = this;
			if (options == 'export') {
				var data = [];
				$instance.find('input.editbox').each(function () {
					var val = $(this).val();
					if (val != null && val != '') data.push({ a: $(this).attr('id'), v: val, f: $(this).attr('data-formula') });
				});
				return data;
			}
		} else {
			options = $.extend({}, $.excelHelpers.defaults, options);
			return this.each(function () {
				var $instance = $(this);
				if ($instance.hasClass('excelsheet') == false) $instance.addClass('excelsheet')
				
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
								var address = $.excelHelpers.getAddressAt(i, j);
								var $td = $('<td>')
									.attr('data-cell-address', address)
									.appendTo($tr);
									
								var $input = $('<input id="' + address + '" type="text" class="editbox" />')
									.focus(function () {
										var parent = $(this).parent();
										$instance.data('activeCellObject', parent);
										$.excelHelpers.setSelection($instance, parent);
									})
									.on('keydown', function (evt) {
										evt = evt ? evt : window.event;
										var charCode = (evt.which) ? evt.which : evt.keyCode;
										if (charCode == 13) {
											// todo: goto next cell
										}
									})
									.appendTo($td);
								
								if (options.data) {
									$.each(options.data, function () {
										if (this.a == address) {
											$input.val(this.v);
											if (this.f && this.f != '') $input.attr('data-formula', this.f)
										}
									})
								}
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
				
			})
		}
	}
})(jQuery)