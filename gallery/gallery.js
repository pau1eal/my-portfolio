const NUM_ROWS = 3;
const DEFAULT_EXCEL = 'proyectos.xlsx';

const PASTEL_PALETTE = {
  mint: {
    bg: '#E0F2F1',
    text: '#004D40'
  },

  rose: {
    bg: '#FCE4EC',
    text: '#880E4F'
  },

  lavender: {
    bg: '#F3E5F5',
    text: '#4A148C'
  },

  azure: {
    bg: '#E1F5FE',
    text: '#01579B'
  },

  gold: {
    bg: '#FFF8E1',
    text: '#8D6E00'
  },

  silver: {
    bg: '#ECEFF1',
    text: '#37474F'
  },

  default: {
    bg: '#ECEFF1',
    text: '#37474F'
  },

  none: {
    bg: '#ECEFF1',
    text: '#37474F'
  }
};

const FALLBACK_COLOR = {
  bg: '#ECEFF1',
  text: '#37474F'
};


const gallery = document.getElementById('gallery');
const fileInput = document.getElementById('file-input');
const tagFilters = document.getElementById('tag-filters');


let rowsData = [];
let positions = [];
let rowWidths = [];

let pausedRows = new Set();

let wheelVelocity = 0;
let animationFrame = null;
let lastTime = performance.now();

let activeTag = 'all';


/* ============================================================
   ROW CONFIGURATION
   ============================================================ */

function generateRowConfig(numRows) {
  return Array.from(
    { length: numRows },
    (_, i) => {

      const isEvenRow = (i + 1) % 2 === 0;

      return {
        isEvenRow,

        minSpeed: isEvenRow
          ? 0.4
          : 0.8,

        direction: isEvenRow
          ? -1
          : 1,

        brickOffset: isEvenRow
      };
    }
  );
}


/* ============================================================
   VALUE CLEANING
   ============================================================ */

function cleanValue(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value).trim();
}


/* ============================================================
   EXCEL HTML VALUE PARSING
   ============================================================ */

function parseSrc(value) {

  const text = cleanValue(value);

  if (!text) {
    return '';
  }

  const match = text.match(
    /src\s*=\s*["']([^"']+)["']/i
  );

  return match
    ? match[1]
    : text;
}


function parseHref(value) {

  const text = cleanValue(value);

  if (!text) {
    return '';
  }

  const match = text.match(
    /href\s*=\s*["']([^"']+)["']/i
  );

  return match
    ? match[1]
    : text;
}


/* ============================================================
   COLOR
   ============================================================ */

function resolveColor(value) {

  const key = cleanValue(value).toLowerCase();

  const aliases = {
    green: 'mint',
    pink: 'rose',
    purple: 'lavender',
    blue: 'azure',
    yellow: 'gold',
    grey: 'silver',
    gray: 'silver'
  };

  const normalizedKey =
    aliases[key] || key;

  return (
    PASTEL_PALETTE[normalizedKey] ||
    FALLBACK_COLOR
  );
}


/* ============================================================
   NORMALIZE EXCEL ROW
   ============================================================ */

function normalizeItem(row) {

  const colorValue =
    cleanValue(
      row['color default'] ??
      row['silver']
    );

  return {

    id:
      cleanValue(row['ID']),

    title:
      cleanValue(row['display-name']),

    tags:
      cleanValue(row['tags']),

    image:
      parseSrc(row['imagen']),

    tier:
      cleanValue(row['tier'])
        .toUpperCase(),

    link:
      parseHref(row['links']),

    colorValue,

    color:
      resolveColor(colorValue)
  };
}


/* ============================================================
   PARAMETER VALIDATION
   ============================================================ */

/*
  A project only enters the gallery if every parameter
  expected by the portfolio is present in Excel.

  This keeps the system procedural.
  Nothing is filtered by project name.
*/

function hasAllParameters(item) {

  return Boolean(

    item.id &&
    item.title &&
    item.tags &&
    item.image &&
    item.tier &&
    item.link &&
    item.colorValue

  );
}


/* ============================================================
   WEIGHTED RANDOMIZATION
   ============================================================ */

function buildWeightedList(items) {

  const weights = {
    S: 5,
    A: 4,
    B: 3,
    C: 2,
    D: 1
  };

  const weighted = [];

  items.forEach(item => {

    const weight =
      weights[item.tier] || 1;

    for (
      let i = 0;
      i < weight;
      i++
    ) {
      weighted.push(item);
    }

  });


  /*
    Fisher-Yates shuffle
  */

  for (
    let i = weighted.length - 1;
    i > 0;
    i--
  ) {

    const j =
      Math.floor(
        Math.random() * (i + 1)
      );

    [
      weighted[i],
      weighted[j]
    ] = [
      weighted[j],
      weighted[i]
    ];
  }


  return weighted;
}


/* ============================================================
   DISTRIBUTE PROJECTS INTO ROWS
   ============================================================ */

function distributeIntoRows(
  items,
  numRows
) {

  const configs =
    generateRowConfig(numRows);


  /*
    Large projects:
    only S and A.

    These are allowed in the larger
    even-numbered rows.
  */

  const largeItems =
    items.filter(
      item =>
        item.tier === 'A' ||
        item.tier === 'S'
    );


  const allItems = [...items];


  const largeRows =
    configs
      .map(
        (config, index) => ({
          config,
          index
        })
      )
      .filter(
        ({ config }) =>
          config.isEvenRow
      )
      .map(
        ({ index }) =>
          index
      );


  const allRows =
    configs
      .map(
        (config, index) => ({
          config,
          index
        })
      )
      .filter(
        ({ config }) =>
          !config.isEvenRow
      )
      .map(
        ({ index }) =>
          index
      );


  const result =
    Array.from(
      { length: numRows },
      () => []
    );


  function distributePool(
    pool,
    targetRows
  ) {

    if (
      !targetRows.length ||
      !pool.length
    ) {
      return;
    }


    const shuffled =
      buildWeightedList(pool);


    shuffled.forEach(
      (item, index) => {

        const rowIndex =
          targetRows[
            index % targetRows.length
          ];

        result[rowIndex].push(item);
      }
    );
  }


  distributePool(
    largeItems,
    largeRows
  );


  distributePool(
    allItems,
    allRows
  );


  return result;
}


/* ============================================================
   EXCEL PARSING
   ============================================================ */

function parseExcel(arrayBuffer) {

  const workbook =
    XLSX.read(
      arrayBuffer,
      { type: 'array' }
    );


  const sheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];


  const rawRows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        defval: null
      }
    );


  const items =
    rawRows
      .map(normalizeItem)
      .filter(hasAllParameters);


  return distributeIntoRows(
    items,
    NUM_ROWS
  );
}


/* ============================================================
   HTML ESCAPING
   ============================================================ */

function escapeHtml(value) {

  return String(value)

    .replaceAll(
      '&',
      '&amp;'
    )

    .replaceAll(
      '<',
      '&lt;'
    )

    .replaceAll(
      '>',
      '&gt;'
    )

    .replaceAll(
      '"',
      '&quot;'
    )

    .replaceAll(
      "'",
      '&#039;'
    );
}


/* ============================================================
   CREATE PROJECT CARD
   ============================================================ */

function createCard(item) {

  const card =
    document.createElement('a');


  card.className =
    'gallery-item';


  card.href =
    item.link &&
    item.link !== 'none'
      ? item.link
      : '#';


  card.dataset.id =
    item.id;


  card.dataset.tier =
    item.tier;


  /*
    External links open in a new tab.
  */

  if (
    item.link &&
    item.link !== 'none'
  ) {

    card.target = '_blank';

    card.rel =
      'noopener noreferrer';

  } else {

    card.addEventListener(
      'click',
      event => {
        event.preventDefault();
      }
    );
  }


  /*
    Tags are read directly from Excel.
  */

  const tags =
    item.tags
      .split(',')
      .map(
        tag => tag.trim()
      )
      .filter(Boolean);


  const tagsHTML =
    tags
      .map(
        tag => `
          <span class="gallery-tag">
            ${escapeHtml(tag)}
          </span>
        `
      )
      .join('');


  card.innerHTML = `
  <img
    src="${escapeHtml(item.image)}"
    alt="${escapeHtml(item.title)}"
    loading="lazy"
  >

  <div class="item-overlay">

    <div class="item-title">
      ${escapeHtml(item.title)}
    </div>

    <div class="gallery-tags">
      ${tagsHTML}
    </div>

  </div>
`;


  /*
    Project color comes from Excel.
  */

  card.style.setProperty(
    '--item-bg',
    item.color.bg
  );


  card.style.setProperty(
    '--item-text',
    item.color.text
  );


  /*
    Pause the row while hovering
    over a project.
  */

  card.addEventListener(
    'mouseenter',
    () => {

      const row =
        card.closest(
          '.gallery-row'
        );

      if (!row) {
        return;
      }

      const index =
        Number(
          row.dataset.rowIndex
        );

      pausedRows.add(index);
    }
  );


  card.addEventListener(
    'mouseleave',
    () => {

      const row =
        card.closest(
          '.gallery-row'
        );

      if (!row) {
        return;
      }

      const index =
        Number(
          row.dataset.rowIndex
        );

      pausedRows.delete(index);
    }
  );


  return card;
}


/* ============================================================
   GET TAGS
   ============================================================ */

function getAllTags(items) {

  const tags =
    new Set();


  items.forEach(item => {

    item.tags
      .split(',')
      .map(
        tag => tag.trim()
      )
      .filter(Boolean)
      .forEach(tag => {

        tags.add(tag);

      });

  });


  return [...tags].sort(
    (a, b) =>
      a.localeCompare(
        b,
        undefined,
        {
          sensitivity: 'base'
        }
      )
  );
}


/* ============================================================
   TAG FILTER BUTTONS
   ============================================================ */

function renderTagFilters(items) {

  if (!tagFilters) {
    return;
  }


  const tags =
    getAllTags(items);


  tagFilters.innerHTML = '';


  const allButton =
    document.createElement('button');


  allButton.type = 'button';

  allButton.className =
    'tag-filter active';

  allButton.dataset.tag =
    'all';

  allButton.textContent =
    'All';


  allButton.addEventListener(
    'click',
    () => {
      setActiveTag('all');
    }
  );


  tagFilters.appendChild(
    allButton
  );


  tags.forEach(tag => {

    const button =
      document.createElement('button');


    button.type = 'button';

    button.className =
      'tag-filter';

    button.dataset.tag =
      tag;

    button.textContent =
      tag;


    button.addEventListener(
      'click',
      () => {
        setActiveTag(tag);
      }
    );


    tagFilters.appendChild(
      button
    );

  });
}


/* ============================================================
   FILTER PROJECTS
   ============================================================ */

function setActiveTag(tag) {

  activeTag = tag;


  /*
    Update active filter button.
  */

  document
    .querySelectorAll(
      '.tag-filter'
    )
    .forEach(button => {

      button.classList.toggle(
        'active',
        button.dataset.tag === tag
      );

    });


  /*
    Instead of simply hiding individual cards,
    we rebuild the visible rows.

    This is important because the infinite
    scrolling calculation needs to know the
    actual width of the visible sequence.
  */

  const filteredItems =
    rowsData
      .flat()
      .filter(item => {

        if (tag === 'all') {
          return true;
        }

        return item.tags
          .split(',')
          .map(
            itemTag =>
              itemTag.trim()
          )
          .includes(tag);

      });


  const newRows =
    distributeIntoRows(
      filteredItems,
      NUM_ROWS
    );


  renderRowsOnly(newRows);
}


/* ============================================================
   RENDER GALLERY
   ============================================================ */

function renderGallery(data) {

  rowsData = data;

  renderTagFilters(
    data.flat()
  );

  renderRowsOnly(data);
}


/* ============================================================
   RENDER ROWS
   ============================================================ */

function renderRowsOnly(data) {

  gallery.innerHTML = '';

  positions =
    data.map(() => 0);

  rowWidths =
    data.map(() => 0);


  const configs =
    generateRowConfig(NUM_ROWS);


  data.forEach(
    (items, rowIndex) => {

      const config =
        configs[rowIndex];


      const row =
        document.createElement('div');


      row.className =
        `gallery-row ${
          config.isEvenRow
            ? 'row-center'
            : 'row-outer'
        }`;


      row.dataset.rowIndex =
        rowIndex;


      const track =
        document.createElement('div');


      track.className =
        'gallery-track';


      /*
        Repeat the same procedural sequence
        several times so the animation has
        enough content to loop seamlessly.
      */

      const repeatCount = 4;


      for (
        let repeat = 0;
        repeat < repeatCount;
        repeat++
      ) {

        items.forEach(item => {

          track.appendChild(
            createCard(item)
          );

        });

      }


      /*
        Even rows have the brick offset.
      */

      if (config.brickOffset) {

        track.style.paddingLeft =
          '3rem';

      }


      row.appendChild(track);

      gallery.appendChild(row);

    }
  );


  requestAnimationFrame(
    measureRows
  );
}


/* ============================================================
   MEASURE LOOP WIDTH
   ============================================================ */

function measureRows() {

  const rowElements =
    [
      ...gallery.querySelectorAll(
        '.gallery-row'
      )
    ];


  rowElements.forEach(
    (row, index) => {

      const track =
        row.querySelector(
          '.gallery-track'
        );


      const itemsPerSet =
        rowsData[index]?.length || 0;


      if (
        !track ||
        !itemsPerSet ||
        !track.children[
          itemsPerSet
        ]
      ) {

        rowWidths[index] = 0;

        return;
      }


      /*
        The position of the first item
        in the second repetition is
        the exact width of one sequence.
      */

      rowWidths[index] =
        track.children[
          itemsPerSet
        ].offsetLeft;

    }
  );
}


/* ============================================================
   ANIMATION
   ============================================================ */

function animate(time) {

  const dt =
    Math.min(
      (time - lastTime) / 16.6667,
      3
    );


  lastTime = time;


  const configs =
    generateRowConfig(
      NUM_ROWS
    );


  positions =
    positions.map(
      (position, index) => {

        if (
          pausedRows.has(index) ||
          !rowWidths[index]
        ) {
          return position;
        }


        const config =
          configs[index];


        let speed =
          config.minSpeed *
          config.direction;


        speed +=
          wheelVelocity *
          config.direction;


        let next =
          position +
          speed *
          dt;


        const width =
          rowWidths[index];


        while (
          next <= -width
        ) {
          next += width;
        }


        while (
          next >= 0
        ) {
          next -= width;
        }


        return next;

      }
    );


  /*
    Wheel momentum gradually disappears.
  */

  wheelVelocity *=
    Math.pow(
      0.9,
      dt
    );


  gallery
    .querySelectorAll(
      '.gallery-row'
    )
    .forEach(
      (row, index) => {

        const track =
          row.querySelector(
            '.gallery-track'
          );


        if (track) {

          track.style.transform =
            `translate3d(
              ${positions[index]}px,
              0,
              0
            )`;

        }

      }
    );


  animationFrame =
    requestAnimationFrame(
      animate
    );
}


/* ============================================================
   START ANIMATION
   ============================================================ */

function startAnimation() {

  if (animationFrame) {

    cancelAnimationFrame(
      animationFrame
    );

  }


  lastTime =
    performance.now();


  animationFrame =
    requestAnimationFrame(
      animate
    );
}


/* ============================================================
   MOUSE WHEEL
   ============================================================ */

gallery.addEventListener(
  'wheel',
  event => {

    event.preventDefault();


    const delta =
      Math.sign(
        event.deltaY
      ) *
      Math.min(
        Math.abs(event.deltaY),
        80
      );


    wheelVelocity +=
      delta * 0.12;

  },
  {
    passive: false
  }
);


/* ============================================================
   WINDOW RESIZE
   ============================================================ */

window.addEventListener(
  'resize',
  () => {
    requestAnimationFrame(
      measureRows
    );
  }
);


/* ============================================================
   MANUAL EXCEL UPLOAD
   ============================================================ */

fileInput?.addEventListener(
  'change',
  event => {

    const file =
      event.target.files?.[0];


    if (!file) {
      return;
    }


    const reader =
      new FileReader();


    reader.onload =
      event => {

        try {

          const data =
            parseExcel(
              event.target.result
            );


          activeTag = 'all';

          renderGallery(data);


        } catch (error) {

          console.error(
            'Could not read Excel file:',
            error
          );


          alert(
            'No se ha podido leer el Excel.'
          );

        }

      };


    reader.readAsArrayBuffer(
      file
    );

  }
);


/* ============================================================
   DEFAULT EXCEL
   ============================================================ */

async function loadDefaultExcel() {

  try {

    const url =
      new URL(
        DEFAULT_EXCEL,
        document.baseURI
      );


    const response =
      await fetch(
        url.href,
        {
          cache: 'no-store'
        }
      );


    if (!response.ok) {

      throw new Error(
        `HTTP ${response.status}`
      );

    }


    const buffer =
      await response.arrayBuffer();


    const data =
      parseExcel(buffer);


    renderGallery(data);


    console.info(
      `Portfolio loaded from ${url.href}`
    );


  } catch (error) {

    console.error(
      'Could not automatically load the default Excel:',
      error
    );


    if (
      window.location.protocol ===
      'file:'
    ) {

      console.warn(
        'Automatic Excel loading requires ' +
        'the portfolio to be served through HTTP/HTTPS.'
      );

    }

  }

}


/* ============================================================
   INIT
   ============================================================ */

loadDefaultExcel();

startAnimation();