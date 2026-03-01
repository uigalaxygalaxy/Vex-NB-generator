<script>
	import { goto } from '$app/navigation';
	import { codeToHtml } from 'shiki';
	import Check from './check.svelte';
	import Number from './number.svelte';

	let pagesToSkip = $state(12);
	let includeDate = $state(true);
	let includeIteration = $state(true);
	let includeColor = $state(true);
	let includeTitle = $state(true);
	let includeLink = $state(true);

	let pageChaining = $state(true);

	let dateColumn = $state(3);
	let iterationColumn = $state(2);
	let titleColumn = $state(1);
	let colorColumn = $state(0);
	let pageNumberColumn = $state(0);

	let pageNumberCoords = $state({
		left: 532,
		top: 0,
		tolerance: {
			left: 21,
			top: 12
		}
	});
	let titleCoords = $state({
		left: 65,
		top: 6,
		tolerance: {
			left: 45,
			top: 22
		}
	});
	let dateCoords = $state({
		left: 511,
		top: 746,
		tolerance: {
			left: 37,
			top: 18
		}
	});
	let iterationCoords = $state({
		left: 52,
		top: 56,
		tolerance: {
			left: 25,
			top: 18
		}
	});
	let ToCCoords = $state({
		left: 43,
		top: 71,
		tolerance: {
			left: 9,
			top: 12
		}
	});

	let ToCDimensions = $state({
		rows: 22,
		columns: 4
	});

	let highlightedCode = $state('');

	let code = $derived(`function notebook() {

    /* HOW TO USE THIS SCRIPT:

    your table of contents MUST BE A TABLE.
    
    make sure your page number text box, title text box, date text box, and iteration text box ARE ACTUALLY TEXT BOXES.

    Please test this on a copy of your notebook first, if ur nb breaks b4 states that's funny and is on you lol

    This script:
    - Finds the page number text box on each slide and updates it to be the correct page number
    - Extracts the title, date, and iteration from each slide based on their coordinates
    - Adds to the table of contents with the page number, title, date, and iteration for each slide, and adds a link to each slide in the title cell

    To find coordinates for each element, create a slide with JUST the element you're tryna to find. Then go to Extensions > App Script > and then run this code: It will read the coordinates of each element and log it.

    const slideShow = SlidesApp.getActivePresentation();
    for (let q = 1; q < slideShow.getSlides().length; q++) {
        let slide = SlidesApp.getActivePresentation().getSlides()[q];
        let elements = slide.getPageElements();
                elements.forEach(element => {
    console.log(\`Checking element on Slide \${q + 1} at (\${element.getLeft()}, \${element.getTop()}) of type \${element.getPageElementType()}\`);
});
}

put whatever it says as ur coordinates.

if ur a sloppy bum increase ur tolerance (please be consistent with your placement, youre gonna be really miserable if you don't)

Default ToC placement is like this:
Page Num (with color) | Title (with link) | Iteration | Date

make sure you have enough table of contents tables before-hand, it will error and stop early but wont do anything bad 

We DO NOT GENERATE ANYTHING. Soo:

The page numbers will ONLY INCREMENT if the Page Number ELEMENT EXISTS. 

Copy and paste A LOT OF YOUR TABLE OF CONTENT PAGES BEFORE HAND.

the script also changes the page number elements btw



*/


    let pagesToSkip = ${pagesToSkip}; // how many pages before the script starts looking for page numbers, title, date, etc. (so you can have a cover page and stuff without it breaking)

    let pageNumberCoords = {
        left: ${pageNumberCoords.left},
        top: ${pageNumberCoords.top},
        tolerance: {
            left: ${pageNumberCoords.tolerance.left},
            top: ${pageNumberCoords.tolerance.top}
        }
    }
    let titleCoords = {
        left: ${titleCoords.left},
        top: ${titleCoords.top},
        tolerance: {
            left: ${titleCoords.tolerance.left},
            top: ${titleCoords.tolerance.top}
        }
    }
    let dateCoords = {
        left: ${dateCoords.left},
        top: ${dateCoords.top},
        tolerance: {
            left: ${dateCoords.tolerance.left},
            top: ${dateCoords.tolerance.top}
        }
    }
    let iterationCoords = {
        left: ${iterationCoords.left},
        top: ${iterationCoords.top},
        tolerance: {
            left: ${iterationCoords.tolerance.left},
            top: ${iterationCoords.tolerance.top}
        }
    }

    let ToCCoords = {
        left: ${ToCCoords.left},
        top: ${ToCCoords.top},
        tolerance: {
            left: ${ToCCoords.tolerance.left},
            top: ${ToCCoords.tolerance.top}
        }
    }

    let ToCDimensions = { //how much row and column in each table of content page
        rows: ${ToCDimensions.rows},
        columns: ${ToCDimensions.columns}
    }

    // If your TOC has different columns, change these. left to right starting at 0
    let pageNumberColumn = ${pageNumberColumn};
    let colorColumn = ${colorColumn};
    let titleColumn = ${titleColumn};
    let iterationColumn = ${iterationColumn};
    let dateColumn = ${dateColumn};

    // disable these if you dont want them in your nb
    let includeDate = ${includeDate};
    let includeIteration = ${includeIteration};
    let includeColor = ${includeColor};
    let includeTitle = ${includeTitle};

    let pageChaining = ${pageChaining};
    //pages with the same titles will have be chained together in the ToC.
    // it will chain it as well if the title includes <cont.>
    //




    // If you DO NOT KNOW WHAT YOU ARE DOING, you should probably not touch this
    // change constants above, but like if ur nb is different ask chat gpt lol
    const slideShow = SlidesApp.getActivePresentation();
    let currentPage = 0;


    tableOfContents = [{
        title: "test slide",
        page: 0,
        date: "1-1-67",
        color: null,
        iteration: "n/a",
        id: '',
        pageStart: 0,
        pageEnd: 0
    }]

    for (let q = pagesToSkip; q < slideShow.getSlides().length; q++) {
        let slide = SlidesApp.getActivePresentation().getSlides()[q];
        if (!slide) break;

        let elements = slide.getPageElements();
        let title = '';
        let date = '';
        let color = null;
        let pageElement = null;
        let iteration = '';
        let slideID = slide.getObjectId();
        let pageStart = null;
        let pageEnd = null;
        elements.forEach(element => {



            if (Math.abs(element.getLeft() - pageNumberCoords.left) < pageNumberCoords.tolerance.left && Math.abs(element.getTop() - pageNumberCoords.top) < pageNumberCoords.tolerance.top) {
                console.log(\`Found page number element on Slide \${q + 1} at (\${element.getLeft()}, \${element.getTop()})\`);
                currentPage++;
                if (element.asShape && element.asShape().getText) {
                    element.asShape().getText().setText(currentPage.toString());
                    let fill = element.asShape().getFill();
                    if (fill.getType() === SlidesApp.FillType.SOLID) {
                        color = fill.getSolidFill().getColor();
                    }
                    pageElement = element;
                }
            }

            if (Math.abs(element.getLeft() - titleCoords.left) < titleCoords.tolerance.left && Math.abs(element.getTop() - titleCoords.top) < titleCoords.tolerance.top) {
                if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                    title = element.asShape().getText().asString();
                }
            }
            if (Math.abs(element.getLeft() - dateCoords.left) < dateCoords.tolerance.left && Math.abs(element.getTop() - dateCoords.top) < dateCoords.tolerance.top) {
                if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                    date = element.asShape().getText().asString();
                }
            }
            if (Math.abs(element.getLeft() - iterationCoords.left) < iterationCoords.tolerance.left && Math.abs(element.getTop() - iterationCoords.top) < iterationCoords.tolerance.top) {
                if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                    iteration = element.asShape().getText().asString();
                }
            }





        });

        console.log(\`Slide \${q + 1}: Title: \${title}, Date: \${date}, ID: \${slideID}, Color: \${color}, Iteration: \${iteration}, Page Element Found: \${!!pageElement}\`);

        if (pageElement) {

            if (!title) title = \`Couldn't find title :c\`;
            if (!date) date = \`Can't find ;-;\`;
            if (!color) color = rgb(255, 255, 255);
            if (!iteration) iteration = \`Can't find :P\`;
            if (!slideID) slideID = \`g3c7d2831742_1_3\`;

            let pastEntry = tableOfContents[tableOfContents.length - 1];
            let pastTitle = tableOfContents[tableOfContents.length - 1]?.title;
            let currentTitle = title;

            let isChain = false;
            if (pageChaining && pastTitle) {
                if (currentTitle === pastTitle || currentTitle === pastTitle.replace(/<cont\.?>/ig, "").trim() || currentTitle.replace(/<cont\.?>/ig, "").trim() === pastTitle || currentTitle.replace(/<cont\.?>/ig, "").trim() === pastTitle.replace(/<cont\.?>/ig, "").trim()) {
                    isChain = true;
                }
            }

            if (isChain) {
                pastEntry.pageEnd = currentPage;
            } else {
                tableOfContents.push({
                    title: title,
                    date: date,
                    color: color,
                    page: currentPage,
                    iteration: iteration,
                    id: slideID,
                    pageStart: currentPage,
                    pageEnd: currentPage
                });
            }


        }


    }
    let pageChainConstructor = '';
    for (let i = 0; i < (Math.ceil(tableOfContents.length / ToCDimensions.rows)); i++) {
        SlidesApp.getActivePresentation().getSlides()[1 + i].getPageElements().forEach(element => {
            if (Math.abs(element.getLeft() - ToCCoords.left) < ToCCoords.tolerance.left && Math.abs(element.getTop() - ToCCoords.top) < ToCCoords.tolerance.top) {
                if (element.asTable) {
                    let table = element.asTable();
                    for (let j = 0; j < ToCDimensions.rows; j++) {
                        let entry = tableOfContents[i * ToCDimensions.rows + j + 1];
                        if (entry) {

                            let pageString = (entry.pageStart === entry.pageEnd)
                                ? entry.pageStart.toString()
                                : \`\${entry.pageStart}-\${entry.pageEnd}\`;

                            table.getCell(j + 1, pageNumberColumn).getText().setText(pageString);

                            // Safety check for the color
                            if (entry.color && includeColor) {
                                // .setSolidFill() accepts both a hex string OR a Color object
                                table.getCell(j + 1, colorColumn).getFill().setSolidFill(entry.color);
                            }

                            if (includeTitle) table.getCell(j + 1, titleColumn).getText().setText(entry.title);
                            if (includeColor) table.getCell(j + 1, titleColumn).getText().getTextStyle().setLinkUrl(\`#slide=id.\${entry.id}\`);
                            if (includeIteration) table.getCell(j + 1, iterationColumn).getText().setText(entry.iteration);
                            if (includeDate) table.getCell(j + 1, dateColumn).getText().setText(entry.date);


                            console.log(\`Page: \${entry.page}, Color: \${entry.color}, Title: \${entry.title}, Date: \${entry.date}\`);

                        }
                    }
                }
            }
        });
    }
}
`);

	async function highlight() {
		highlightedCode = await codeToHtml(code, {
			lang: 'javascript',
			theme: 'gruvbox-dark-medium'
		});
	}

	$effect(() => {
		pagesToSkip;
		includeDate;
		includeIteration;
		includeColor;
		includeTitle;
		pageChaining;
		dateColumn;
		iterationColumn;
		titleColumn;
		colorColumn;
		pageNumberColumn;
		dateCoords;
		iterationCoords;
		titleCoords;
		pageNumberCoords;
		ToCCoords;
		ToCDimensions;
		highlight();
	});

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

<h1 class="subHeading ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">
	How this will work
</h1>
<h1 class="mt-1 ml-4 font-[D-Din] text-lg text-[#ffe5b5]">
	Answer a series of questions about your notebook, and the script will change the constants (as you
	can see at the top of the script) accordingly! This is meant for non-programmers, since code is
	super scary.
</h1>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		How much pages should the script skip (because of cover pages, and whatnot)?
	</h1>
	<Number bind:value={pagesToSkip} text="pages to skip" />
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		What elements do you want to be tracked?
	</h1>
	<Check bind:checked={includeDate} text="Track the Date?" />
	<Check bind:checked={includeIteration} text="Track the Iteration/Subheading?" />
	<Check bind:checked={includeColor} text="Track the Color of the Page Num?" />
	<Check bind:checked={includeLink} text="Track the Link?" />
	<Check bind:checked={includeTitle} text="Track the Title? (not sure why you wouldn't)" />
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		Do you want page chaining?
	</h1>
	<Check
		bind:checked={pageChaining}
		text=" Page Chaining: in the table of contents they will be in the same row with a combined page number (i.e. like 41-67) "
	/>
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		What column do you want each element to be in your table of contents? (starting at 0, left to
		right)
	</h1>
	{#if includeDate}<Number bind:value={dateColumn} text="column, Date" />
	{/if}
	{#if includeIteration}<Number bind:value={iterationColumn} text="column, Iteration/Subheading" />
	{/if}
	{#if includeTitle}<Number bind:value={titleColumn} text="column, Title" />
	{/if}
	{#if includeColor}<Number
			bind:value={colorColumn}
			text="column, What column the color is applied to"
		/>
	{/if}
	<Number bind:value={pageNumberColumn} text="column, Page Number" />
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		What coordinates are your elements?
	</h1>
	{#if includeDate}
		<div class="mt-2 flex flex-row gap-2">
			<Number bind:value={dateCoords.left} text=", " />
			<Number bind:value={dateCoords.top} text="X (from left), Y (from top) - Date Element" />
		</div>
		<div class="flex flex-row gap-2">
			<span class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">±</span>
			<Number bind:value={dateCoords.tolerance.left} text=", " />
			<Number
				bind:value={dateCoords.tolerance.top}
				text="X (from left), Y (from top) - Element Tolerance"
			/>
		</div>
	{/if}

	{#if includeIteration}
		<div class="mt-2 flex flex-row gap-2">
			<Number bind:value={iterationCoords.left} text=", " />
			<Number
				bind:value={iterationCoords.top}
				text="X (from left), Y (from top) - Iteration/Subheading Element"
			/>
		</div>
		<div class="flex flex-row gap-2">
			<span class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">±</span>
			<Number bind:value={iterationCoords.tolerance.left} text=", " />
			<Number
				bind:value={iterationCoords.tolerance.top}
				text="X (from left), Y (from top) - Element Tolerance"
			/>
		</div>
	{/if}

	{#if includeTitle}
		<div class="mt-2 flex flex-row gap-2">
			<Number bind:value={titleCoords.left} text=", " />
			<Number bind:value={titleCoords.top} text="X (from left), Y (from top) - Title Element" />
		</div>
		<div class="flex flex-row gap-2">
			<span class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">±</span>
			<Number bind:value={titleCoords.tolerance.left} text=", " />
			<Number
				bind:value={titleCoords.tolerance.top}
				text="X (from left), Y (from top) - Element Tolerance"
			/>
		</div>
	{/if}

	<div class="mt-2 flex flex-row gap-2">
		<Number bind:value={pageNumberCoords.left} text=", " />
		<Number
			bind:value={pageNumberCoords.top}
			text="X (from left), Y (from top) - Page Number Element"
		/>
	</div>
	<div class="flex flex-row gap-2">
		<span class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">±</span>
		<Number bind:value={pageNumberCoords.tolerance.left} text=", " />
		<Number
			bind:value={pageNumberCoords.tolerance.top}
			text="X (from left), Y (from top) - Element Tolerance"
		/>
	</div>
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		How much rows does your table of contents have?
	</h1>
	<Number bind:value={ToCDimensions.rows} text="rows" />
</div>

<div class="align-front justify-left mt-8 ml-4 flex w-screen flex-col">
	<h1 class="otherTitle mt-1 font-[D-Din] text-xl font-extrabold text-transparent">
		What are the coordinates of your table of contents?
	</h1>
	<div class="mt-2 flex flex-row gap-2">
		<Number bind:value={ToCCoords.left} text=", " />
		<Number
			bind:value={ToCCoords.top}
			text="X (from left), Y (from top) - Table of Contents Element"
		/>
	</div>
	<div class="flex flex-row gap-2">
		<span class="mt-1 ml-2 font-[D-Din] text-lg text-[#ffe5b5]">±</span>
		<Number bind:value={ToCCoords.tolerance.left} text=", " />
		<Number
			bind:value={ToCCoords.tolerance.top}
			text="X (from left), Y (from top) - Element Tolerance"
		/>
	</div>
</div>

<h1 class="subHeading mt-8 ml-2 w-70 pl-1 font-[D-Din] text-3xl font-extrabold text-black">
	Generated Script:
</h1>
<div class=" mt-2 ml-2 border-2 border-[#522f01] p-1 text-sm">
	{@html highlightedCode}
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
