<script>
	import { onMount } from "svelte";

	function dateForm(value, option) {
		if (typeof value === "string") {
			let [day, month, year] = value.split(".");
			value = new Date(Number(year) - 543, Number(month) - 1, Number(day));
			return new Date(value).toLocaleDateString("th", {
				day: "numeric",
				month: "long",
				year: "numeric",
				...option,
			});
		}
		return value;
	}
	function moneyForm(value, option) {
		if (isNaN(value)) {
			return value;
		}
		value = Number(value);
		return value == 0
			? ""
			: value.toLocaleString("th", {
					minimumFractionDigits: 2,
					maximumFractionDigits: 2,
					...option,
				});
	}
	function numerize(value) {
		if (typeof value === "string") {
			value = value.trim().replace(/,/g, "");
			return Number(value);
		}
	}

	async function upload(e) {
		const file = e.currentTarget.files[0];
		await file.arrayBuffer().then((raw) => {
			const spreadsheet = XLSX.read(raw, { cellDates: true });
			const sheetName = spreadsheet.SheetNames[0];
			const worksheet = spreadsheet.Sheets[sheetName];
			const rest = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
			let result = [];
			rest.forEach((cols) => {
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
					]);
				}
			});
			result.shift();
			aoa = result;
		});
	}

	let aoa = $state([
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
		[],
	]);
	let user = $state({doc:'ขบ 02', placeCode: '', costCode: '', phone: '', sender: '', approver: ''})

	let calculate = $derived.by(() => {
		let total = [0, 0, 0, 0];
		aoa.forEach((cols) => {
			total[0] += numerize(cols[8]);
			total[1] += numerize(cols[9]);
			total[2] += numerize(cols[10]);
			total[3] += numerize(cols[12]);
		});
		return { total };
	});

	onMount(() => {});
</script>

<div class="p-4 flex flex-wrap gap-4 print:hidden select-none">
	<div class="">
		<label class="cursor-pointer">
			Upload Excel file.xlsx รายงานสรุปรายการเบิกจ่ายของหน่วยงาน (เงินงบ) รายวัน
			<input
				type="file"
				class="cursor-pointer text-cyan-500"
				accept=".xlsx"
				onchange={(e) => {
					upload(e);
				}}
			/>
		</label>
	</div>
	<div class="">
		<button
			class="cursor-pointer text-cyan-500"
			onclick={() => {
				print();
			}}
		>
			Print
		</button>
	</div>
</div>

<div class="mx-auto max-w-fit">
	<table>
		<thead>
			<tr>
				<td class="text-center" colspan="99">
					ใบขออนุมัติหรือยกเลิกการเบิกเงิน (<span></span>)
				</td>
			</tr>
			<tr>
				<td class="text-center" colspan="99">
					ชื่อหน่วยงาน
					<span>{aoa[0][4]}</span>
					รหัสกรม
					<span></span>
					รหัสศูนย์ต้นทุน
					<span></span>
					โทรศัพท์
					<span></span>
				</td>
			</tr>
			<tr>
				<td class="text-center" colspan="99">
					ส่งเข้าระบบวันที่
					<span>{dateForm(aoa[0][0])}</span>
					วันที่ผ่านรายการ
					<span>{dateForm(aoa[0][1])}</span>
				</td>
			</tr>
			<tr>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">ลำดับ<br />ที่</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">เลข GF<br />(10 หลัก)</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">จำนวนเงิน<br />ขอเบิก</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">ภาษี</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">ค่าปรับ</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">จำนวนเงิน<br />ขอรับ</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					colspan="2">ประเภทเอกสาร</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">ให้<br />อนุมัติ</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">ให้<br />ยกเลิก</td
				>
				<td
					class="border-r border-l border-b border-t text-center px-1"
					rowspan="2">หมายเหตุ</td
				>
			</tr>
			<tr>
				<td class="border-r border-l border-b border-t text-center px-1"
					>ประจำเดือน</td
				>
				<td class="border-r border-l border-b border-t text-center px-1"
					>ประจำวัน</td
				>
			</tr>
		</thead>
		<tbody>
			{#each aoa as cols, index}
				<tr>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;">{index + 1}</td
					>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;">{cols[3]}</td
					>
					<td
						class="border-r border-l px-1 text-right"
						style="border-bottom: 1px dotted;">{moneyForm(cols[8])}</td
					>
					<td
						class="border-r border-l px-1 text-right"
						style="border-bottom: 1px dotted;">{moneyForm(cols[9])}</td
					>
					<td
						class="border-r border-l px-1 text-right"
						style="border-bottom: 1px dotted;">{moneyForm(cols[10])}</td
					>
					<td
						class="border-r border-l px-1 text-right"
						style="border-bottom: 1px dotted;">{moneyForm(cols[12])}</td
					>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;"
					></td>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;">{cols[2]}</td
					>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;">/</td
					>
					<td
						class="border-r border-l px-1 text-center"
						style="border-bottom: 1px dotted;"
					></td>
					<td class="border-r border-l px-1" style="border-bottom: 1px dotted;"
						>{cols[6]?.slice(-3) + "/" + cols[6]?.slice(1, 3)}</td
					>
				</tr>
			{/each}
			<tr>
				<td class="border-t text-center" colspan="2">รวมเป็นเงิน</td>
				<td class="border-r border-l border-b border-t px-1 text-right"
					>{moneyForm(calculate.total[0])}</td
				>
				<td class="border-r border-l border-b border-t px-1 text-right"
					>{moneyForm(calculate.total[1])}</td
				>
				<td class="border-r border-l border-b border-t px-1 text-right"
					>{moneyForm(calculate.total[2])}</td
				>
				<td class="border-r border-l border-b border-t px-1 text-right"
					>{moneyForm(calculate.total[3])}</td
				>
				<td class="border-t"></td>
				<td class="border-t"></td>
				<td class="border-t"></td>
				<td class="border-t"></td>
				<td class="border-t"></td>
			</tr>
			<tr>
				<td class="" colspan="99"
					>รวม <span class="">{aoa.length}</span> ฉบับ</td
				>
			</tr>
			<tr>
				<td class="" colspan="99">ลงชื่อ</td>
			</tr>
		</tbody>
	</table>
</div>
