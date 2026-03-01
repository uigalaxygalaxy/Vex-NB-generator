<script>
	let { checked = $bindable(), text } = $props();
</script>

<div class="align-left justify-top mt-1 flex flex-row">
	<input
		bind:checked
		class="color-[#27221e] w-10 bg-[#ffe5b5] px-1 font-[D-Din] text-lg"
		type="checkbox"
	/>
	<h1 class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">{text}</h1>
</div>
