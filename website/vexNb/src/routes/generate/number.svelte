<script>
	let { value = $bindable(), text } = $props();
</script>

<div class="align-left justify-top mt-1 flex h-9 flex-row">
	<input
		bind:value
		class="color-[#27221e] w-20 bg-[#ffe5b5] px-1 font-[D-Din] text-lg"
		type="number"
	/>
	<h1 class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">{text}</h1>
</div>
