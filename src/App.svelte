<script>

		function dateForm (value, option) {
			return new Date(value).toLocaleDateString('th', {
				day: 'numeric',
				month: 'short',
				year: 'numeric',
				...option,
			})
		}
		function moneyForm (value) {
			value = Number(value)
			return value == 0 ? '' : value.toLocaleString('th', {
				minimumFractionDigits: 2,
				maximumFractionDigits: 2,
				...option,
			})
		}

	let	aoa= $state([[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],])
		async function upload (e) {
			const file = e.currentTarget.files[0]
			await file.arrayBuffer().then( function (raw) {
				const spreadsheet = XLSX.read(raw, { cellDates: true })
				const sheetName = spreadsheet.SheetNames[0]
				const worksheet = spreadsheet.Sheets[sheetName]
				const rest = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
				let result = []
				rest.forEach( function (cols) {
					if (cols[3]) {
						result.push([
							cols[1], 
							cols[3], 
							cols[4], 
							cols[6], 
							cols[10], 
							cols[11], 
							cols[12], 
							cols[14], 
							cols[16], 
							cols[19], 
							cols[24], 
							cols[26], 
							cols[29], 
						])
					}
				})
				result.shift()
				this.aoa = result
			})
		}
		
</script>



		<div class="p-4 flex flex-wrap gap-4 print:hidden select-none">
			<div class="">
				<label class="cursor-pointer">
					Upload Excel file.xlsx รายงานสรุปรายการเบิกจ่ายของหน่วยงาน (เงินงบ) รายวัน
					<input type="file" class="cursor-pointer text-cyan-500" accept=".xlsx" @change="upload($event)" />
				</label>
			</div>
			<div class="">
				<button class="cursor-pointer text-cyan-500" onclick="print()">
					Print
				</button>
			</div>
		</div>

		<div class="mx-auto max-w-fit">
			<table>
				<thead>
					<tr>
						<td class="text-center" colspan="99">
							ใบขออนุมัติหรือยกเลิกการเบิกเงิน (<span x-text="''"></span>)

						</td>
					</tr>
					<tr>
						<td class="text-center" colspan="99">
							ชื่อหน่วยงาน
							<span x-text="aoa[0][4]"></span>
							รหัสกรม
							<span x-text="''"></span>
							รหัสศูนย์ต้นทุน
							<span x-text="''"></span>
							โทรศัพท์
							<span x-text="''"></span>
						</td>
					</tr>
					<tr>
						<td class="text-center" colspan="99">
							ส่งเข้าระบบวันที่ เดือน
							<span x-text="aoa[0][0]"></span>
							วันที่ผ่านรายการ เดือน
							<span x-text="aoa[0][1]"></span>
						</td>
					</tr>
					<tr>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">ลำดับ<br>ที่</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">เลข GF<br>(10 หลัก)</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">จำนวนเงิน<br>ขอเบิก</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">ภาษี</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">ค่าปรับ</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">จำนวนเงิน<br>ขอรับ</td>
						<td class="border-r border-l border-b border-t text-center px-1" colspan="2">ประเภทเอกสาร</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">ให้<br>อนุมัติ</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">ให้<br>ยกเลิก</td>
						<td class="border-r border-l border-b border-t text-center px-1" rowspan="2">หมายเหตุ</td>
					</tr>
					<tr>
						<td class="border-r border-l border-b border-t text-center px-1">ประจำเดือน</td>
						<td class="border-r border-l border-b border-t text-center px-1">ประจำวัน</td>
					</tr>
				</thead>
				<tbody>
					<template x-for="(cols, index) in aoa">
						<tr>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text="index+1"></td>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text=" cols[3]"></td>
							<td class="border-r border-l px-1 text-right" style="border-bottom: 1px dotted;"
								x-text="moneyForm(cols[8])"></td>
							<td class="border-r border-l px-1 text-right" style="border-bottom: 1px dotted;"
								x-text="moneyForm(cols[9])"></td>
							<td class="border-r border-l px-1 text-right" style="border-bottom: 1px dotted;"
								x-text="moneyForm(cols[10])"></td>
							<td class="border-r border-l px-1 text-right" style="border-bottom: 1px dotted;"
								x-text="moneyForm(cols[12])"></td>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text="cols[2]"></td>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text="''"></td>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text="'/'"></td>
							<td class="border-r border-l px-1 text-center" style="border-bottom: 1px dotted;" x-text="''"></td>
							<td class="border-r border-l px-1" style="border-bottom: 1px dotted;"
								x-text="cols[6].slice(-3) + '/' + cols[6].slice(1, 3)"></td>
						</tr>
					</template>
					<tr>
						<td class="border-t text-center" colspan="2">รวมเป็นเงิน</td>
						<td class="border-r border-l border-b border-t px-1 text-right" x-text="moneyForm(cols[8])"></td>
						<td class="border-r border-l border-b border-t px-1 text-right" x-text="moneyForm(cols[9])"></td>
						<td class="border-r border-l border-b border-t px-1 text-right" x-text="moneyForm(cols[10])"></td>
						<td class="border-r border-l border-b border-t px-1 text-right" x-text="moneyForm(cols[12])"></td>
						<td class="border-t"></td>
						<td class="border-t"></td>
						<td class="border-t"></td>
						<td class="border-t"></td>
						<td class="border-t"></td>
					</tr>
					<tr>
						<td class="" colspan="99">รวม <span class="" x-text="aoa.length"></span> ฉบับ</td>
					</tr>
					<tr>
						<td class="" colspan="99">ลงชื่อ</td>
					</tr>
				</tbody>
			</table>
		</div>

