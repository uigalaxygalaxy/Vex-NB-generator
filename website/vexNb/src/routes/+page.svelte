<script>
	import { goto } from '$app/navigation';
	import { codeToHtml } from 'shiki';

	let highlightedCode = '';

	let code = `const slideShow = SlidesApp.getActivePresentation();
    for ( let q = 1; q < slideShow.getSlides().length; q++ ) {
        let slide = SlidesApp.getActivePresentation().getSlides()[ q ];
        let elements = slide.getPageElements();
        elements.forEach( element => {
            console.log( \`Checking element on Slide \${q + 1} at (\${element.getLeft()}, \${element.getTop()}) of type \${element.getPageElementType()}\`);
        });
}
`;

	async function highlight() {
		highlightedCode = await codeToHtml(code, {
			lang: 'javascript',
			theme: 'gruvbox-dark-medium'
			// Add line numbers
		});
	}

	highlight();
</script>

<div
	class="mb-4 flex w-screen justify-center"
	onclick={() => goto('/')}
	aria-label="Go back to home page"
	role="button"
	tabindex="0"
	onkeypress={(e) => {
		if (e.key === 'Enter') goto('/');
	}}
>
	<div class="justify-left ml-1 flex w-screen flex-col">
		<p class="title z-4 mt-1 pb-2 font-[Industry] text-6xl text-transparent">Vex Notebook Helper</p>

		<p class="mr-1.1 titleOutline absolute -z-1 mt-3 pb-2 font-[Industry] text-6xl">
			Vex Notebook Helper
		</p>
		<p class="mr-1.1 titleOutline absolute -z-1 pb-2 font-[Industry] text-6xl">
			Vex Notebook Helper
		</p>

		<div class="align-left justify-top -mt-1 ml-1">
			<p class="otherTitle z-4 -mt-0.5 pb-2 font-[D-Din] text-3xl text-transparent">
				for Google Slides
			</p>
		</div>
	</div>
	<img class="logo" src="notebookhelperlogo.png" width="125" alt="Vex Notebook Helper Logo" />
</div>

<h1 class="subHeading ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">Welcome!</h1>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	Vex Notebook Helper is an AppScript that automatically increments page numbers and generates Table
	of Contents for Vex Notebooks on Google Slides!
</h1>

<h1 class="subHeading mt-8 ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">
	What it does
</h1>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	This script will:
	<ul class="list-disc pl-8">
		<li>Find page number elements and increment them</li>
		<li>Extract data from each slide</li>
		<li>Populate a Table of Contents table with this data</li>
	</ul>
</h1>

<h1 class="subHeading mt-8 ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">
	Prerequisites
</h1>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	Vex Notebook Helper is only designed to work on notebooks that:
	<ul class="list-disc pl-8">
		<li>Have the elements you want to track be in a consistent position</li>
		<li>Have a table of contents page (must be a table)</li>
		<li>Must be on Google Slides (duh)</li>
	</ul>
	If your notebook doesn't meet these standards, well sorry :c
</h1>

<h1 class="subHeading mt-8 ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">
	How to Use
</h1>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	Choose to generate your script (coming soon), or download the code and adjust the constants at the
	top of the script! <br />

	Run the script by going to Extensions > Apps Script, pasting the code in, and running the notebook
	function. <br />

	In order to find the coordinates for each element, create a slide with JUST that element, and use
	this code in Apps Script to log the coordinates:
</h1>
<div class=" mt-2 ml-2 border-2 border-[#522f01] p-1 text-sm">
	{@html highlightedCode}
</div>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	It will log the coordinates of EACH element on EVERY slide, so i'd reccomend moving your test
	slides to the top and ending after your test slides have been checked. <br />

	You can then put these coordinates in the script! The tolerance is the amount of wiggle room
	allowed for each element. I'd reccomend around 20.
</h1>
<div class="flex w-screen justify-center">
	<div class="mb-16 flex w-1/2 justify-around">
		<div class="flex w-1/2 flex-col items-center">
			<button
				onclick={() => goto('/code')}
				class="codeButton text-bold mt-4 bg-[#ffe5b5] px-2 py-1 font-[D-Din] text-xl font-extrabold"
			>
				View/Download Script
			</button>
		</div>
		<div class="flex w-1/2 flex-col items-center">
			<button
				class="codeButton text-bold mt-4 bg-[#ffe5b5] px-2 py-1 font-[D-Din] text-xl font-extrabold"
			>
				Generate Script (coming soon)
			</button>
		</div>
	</div>
</div>
<footer>
	<a
		href="https://uigalaxy.net"
		target="_blank"
		rel="noopener noreferrer"
		class="watermark otherTitle mt-8 ml-2 w-full pl-1 text-center font-[D-Din] text-3xl font-extrabold text-transparent underline transition-all duration-300 hover:text-[32px] active:text-[30px]"
	>
		uigalaxy.net | 45434X Paradox
	</a>
	<h1 class="align-center mt-1 w-full text-center font-[D-Din] text-lg text-[#ffe5b5]">
		check out my other projects <br />
		<span class="underline transition-all duration-300 hover:text-white"
			><a href="https://pid-tuner-gamma.vercel.app/" target="_blank" rel="noopener noreferrer"
				>PID Tuner</a
			></span
		>
	</h1>
</footer>
