/** The ONE length grammar every size/length flag parses through — page
 *  `--size`/`--margins`, paragraph spacing/indents, image `--width`/`--height`,
 *  table `--widths`/`--table-width`/`--width`, `--row-height`, `--cell-margins`,
 *  and `--tabs` stop positions. A value is a number with an optional unit suffix
 *  from `LENGTH_UNIT_TWIPS` (in, ", cm, mm, pt, tw/twip/twips/dxa; px only where
 *  the flag opts in). What a BARE number means is the per-flag `bare` option, so
 *  each flag keeps the meaning weak agents already rely on; a suffix always means
 *  what it says; anything else is a `{ error, hint }` (a usage error), never a
 *  silently truncated `parseFloat`/`parseInt` (`--width 3cm` used to be 3 INCHES).
 *  Returns twips (rounded). */
export function parseLength(
	raw: string,
	flag: string,
	options: LengthOptions,
): number | LengthError {
	return parseLengthToken(raw.trim(), flag, options, undefined);
}

/** Parse a comma-separated list of lengths (`--margins`, `--widths`) into twips,
 *  one per entry. A unit on ONLY the last entry, every other entry bare, is the
 *  list's shared unit — the shape `read` prints (`widths="2.375,4.618in"`,
 *  `margins="1,1,1,1.5in"`), so a printed hint re-applies verbatim, and
 *  `2.5,2.5,2.5,2.5cm` means centimeters throughout rather than three bare
 *  inches and one cm. Otherwise each entry carries its own unit or the flag's
 *  bare default. */
export function parseLengthList(
	raw: string,
	flag: string,
	options: LengthOptions,
): number[] | LengthError {
	const parts = raw.split(",").map((part) => part.trim());
	const sharedUnit = sharedTrailingUnit(parts, options);
	const twips: number[] = [];
	for (const part of parts) {
		const value = parseLengthToken(part, flag, options, sharedUnit);
		if (typeof value !== "number") return value;
		// A bare entry that only became huge by borrowing the trailing unit
		// (`1440,2880,1in` → 1440 INCHES) meant the flag's own bare unit, not
		// the shared one — refuse rather than write a 100-foot column/margin.
		if (sharedUnit !== undefined && Math.abs(value) > MAX_SHARED_TWIPS)
			return mixedListError(raw, flag, sharedUnit);
		twips.push(value);
	}
	return twips;
}

/** The unit suffix of the last entry when it is the ONLY suffixed entry in a
 *  multi-entry list and a unit this flag accepts, else undefined (see
 *  `parseLengthList`) — an invalid trailing unit (`1,2,3em`) is left on its own
 *  entry so the error names that entry, not the first bare one. */
function sharedTrailingUnit(
	parts: string[],
	options: LengthOptions,
): string | undefined {
	if (parts.length < 2) return undefined;
	const units = parts.map((part) => part.match(LENGTH_PATTERN)?.[2]);
	const last = units[units.length - 1]?.toLowerCase();
	if (last === undefined || unitFactor(last, options) === undefined)
		return undefined;
	return units.slice(0, -1).every((unit) => unit === undefined)
		? last
		: undefined;
}

function mixedListError(raw: string, flag: string, unit: string): LengthError {
	return {
		error: `Invalid --${flag}: ${raw}`,
		hint: `A unit on only the last value applies to every value, so the bare values here are read as ${unit} — far larger than any page. Give each value its own unit (e.g. 1440tw,2880tw,1in) or use one unit throughout.`,
	};
}

/** One entry → twips, or an error. `sharedUnit` stands in for a missing suffix. */
function parseLengthToken(
	token: string,
	flag: string,
	options: LengthOptions,
	sharedUnit: string | undefined,
): number | LengthError {
	const match = token.match(LENGTH_PATTERN);
	if (!match) return lengthError(token, flag, options);
	const number = Number.parseFloat(match[1] as string);
	const unit = match[2]?.toLowerCase() ?? sharedUnit;
	if (number < 0 && !options.signed) return lengthError(token, flag, options);
	const twips = toTwips(number, unit, token, flag, options);
	// Checked AFTER rounding: `0.01pt` rounds to a 0-twip row/width/image.
	if (twips === 0 && options.positive) return lengthError(token, flag, options);
	return twips;
}

function toTwips(
	number: number,
	unit: string | undefined,
	token: string,
	flag: string,
	options: LengthOptions,
): number | LengthError {
	if (unit === undefined) return bareToTwips(number, token, flag, options);
	const factor = unitFactor(unit, options);
	if (factor === undefined) return lengthError(token, flag, options);
	return Math.round(number * factor);
}

/** Twips per `unit`, honoring the flag's `pixels` opt-in. Own-property lookup
 *  only: a plain-object index would resolve `constructor` through the prototype
 *  and write `w:left="NaN"`. */
function unitFactor(unit: string, options: LengthOptions): number | undefined {
	if (unit === "px") return options.pixels ? TWIPS_PER_PIXEL : undefined;
	return Object.hasOwn(LENGTH_UNIT_TWIPS, unit)
		? LENGTH_UNIT_TWIPS[unit]
		: undefined;
}

/** A bare number in the flag's default unit. Zero is zero in any unit, so it is
 *  accepted even where a unit is otherwise required. Two guards catch a value
 *  copied across the inch/twips boundary: an inch-default flag with
 *  `largeBareIsTwips` reads a bare number above `BARE_INCH_MAX` as twips (no
 *  indent or margin is 10+ inches, but 720/1440 are what `read --ast` reports),
 *  and a twips-default flag REJECTS a bare 1–`BARE_INCH_MAX` (a 2-twip column is
 *  0.001in — that value meant inches; `2in` or `2tw` says which). Bare twips
 *  must be whole numbers: `2.375` is an inch value missing its suffix. */
function bareToTwips(
	number: number,
	token: string,
	flag: string,
	options: LengthOptions,
): number | LengthError {
	if (number === 0) return 0;
	const bare = options.bare;
	if (bare === undefined) return lengthError(token, flag, options);
	if (bare === "tw") {
		if (!Number.isInteger(number) || Math.abs(number) <= BARE_INCH_MAX)
			return ambiguousTwipsError(token, flag);
		return number;
	}
	if (
		bare === "in" &&
		options.largeBareIsTwips &&
		Math.abs(number) > BARE_INCH_MAX
	) {
		return Math.round(number);
	}
	return Math.round(number * (LENGTH_UNIT_TWIPS[bare] as number));
}

function lengthError(
	raw: string,
	flag: string,
	options: LengthOptions,
): LengthError {
	const units = options.pixels
		? "in, cm, mm, pt, px, tw"
		: "in, cm, mm, pt, tw";
	const bare = options.bare;
	const bareRule =
		bare === undefined
			? "A unit is required (a bare 0 is allowed)"
			: bare === "tw"
				? `A bare number is whole twips above ${BARE_INCH_MAX} (1440 = 1in)`
				: `A bare number is ${bare === "in" ? "inches" : "points"}`;
	const sign = options.signed ? "" : " Negative values are not allowed.";
	return {
		error: `Invalid --${flag}: ${raw}`,
		hint: `${bareRule}; or add a unit suffix (${units}), e.g. ${EXAMPLES[bare ?? "required"]}${options.pixels ? ", 300px" : ""}.${sign}`,
	};
}

/** A bare twips value that is really an inch value missing its suffix — name
 *  the likely intent instead of the generic grammar. */
function ambiguousTwipsError(token: string, flag: string): LengthError {
	return {
		error: `Invalid --${flag}: ${token}`,
		hint: `A bare number here is TWIPS (1440 = 1in), so ${token} would be ${token}/1440 of an inch. Add the unit you mean: ${token}in, ${token}cm, ${token}pt (or ${token}tw for twips; bare twips are whole numbers above ${BARE_INCH_MAX}).`,
	};
}

const EXAMPLES: Readonly<Record<NonNullable<BareUnit> | "required", string>> = {
	in: "0.5, 0.5in, 1.27cm, 36pt, 720tw",
	pt: "6, 6pt, 0.1in, 2mm, 120tw",
	tw: "1440, 1in, 2.54cm, 72pt",
	required: "0.4in, 28pt, 1cm, 5mm, 576tw",
};

/** A number (optionally signed, optionally fractional) and an optional unit. */
const LENGTH_PATTERN = /^(-?(?:\d+(?:\.\d*)?|\.\d+))\s*("|[a-z]+)?$/i;

const TWIPS_PER_INCH = 1440;
const TWIPS_PER_POINT = 20;
/** CSS/Word reference pixel: 96 per inch, so exactly 15 twips. */
const TWIPS_PER_PIXEL = TWIPS_PER_INCH / 96;

/** Twips per unit suffix. `tw`/`twips`/`dxa` are the raw OOXML unit — what
 *  `read --ast` reports, so an agent can feed an AST value straight back with
 *  its unit named. `px` is NOT here: it's an opt-in (`pixels`) because only an
 *  image has a natural pixel size. */
const LENGTH_UNIT_TWIPS: Readonly<Record<string, number>> = {
	in: TWIPS_PER_INCH,
	'"': TWIPS_PER_INCH,
	cm: TWIPS_PER_INCH / 2.54,
	mm: TWIPS_PER_INCH / 25.4,
	pt: TWIPS_PER_POINT,
	tw: 1,
	twip: 1,
	twips: 1,
	dxa: 1,
};

/** The inch/twips boundary for bare numbers (see `bareToTwips`). No paragraph
 *  indent or page margin is 10+ inches on any real page, but 119 / 720 / 1440
 *  are exactly the twips values `read --ast` reports — a weak agent copied one
 *  into an inch flag and got a 119-INCH indent that pushed three résumé entries
 *  off the page, exit 0. */
const BARE_INCH_MAX = 9;

/** 22in — Word's largest page side. A list entry that borrowed the trailing
 *  unit and lands past it was a bare value in the flag's own unit. */
const MAX_SHARED_TWIPS = 22 * TWIPS_PER_INCH;

type BareUnit = "in" | "pt" | "tw" | undefined;

/** Per-flag knobs for the shared grammar. */
export type LengthOptions = {
	/** Unit of a bare number; omit to REQUIRE a unit (`--row-height 28` must not
	 *  become a 28-inch row). */
	bare?: BareUnit;
	/** Admit a leading `-` (`w:left`/`w:firstLine`/page margins are signed). */
	signed?: boolean;
	/** Reject zero (a width or size of 0 is never meaningful). */
	positive?: boolean;
	/** Admit `px` at 96 dpi — image sizes only. */
	pixels?: boolean;
	/** Inch-default only: a bare number above `BARE_INCH_MAX` is twips. */
	largeBareIsTwips?: boolean;
};

export type LengthError = { error: string; hint: string };
