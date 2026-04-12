import { DateTime, Duration } from 'luxon';
import { randomUUID } from 'node:crypto';

import type {
	IDataObject,
	IExecuteFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	INodeListSearchResult,
} from 'n8n-workflow';

const DAV_NAMESPACE = 'DAV:';
const CALDAV_NAMESPACE = 'urn:ietf:params:xml:ns:caldav';
const CALENDAR_SERVER_NAMESPACE = 'http://calendarserver.org/ns/';
const PROD_ID = '-//n8n//CalDAV Interface//EN';
const TEXT_MATCH_COLLATION = 'i;unicode-casemap';
const DEFAULT_EVENT_DURATION_MINUTES = 60;
const DEFAULT_RANGE_LIMIT = 50;
const DEFAULT_TIMEZONE = 'UTC';
const DAV_PROPERTY_NAMES = new Set([
	'displayname',
	'resourcetype',
	'current-user-principal',
	'getetag',
	'href',
]);
const DAYS_OF_WEEK = [
	'monday',
	'tuesday',
	'wednesday',
	'thursday',
	'friday',
	'saturday',
	'sunday',
] as const;
const ORDINAL_SUFFIX_PATTERN = /(\d)(st|nd|rd|th)\b/gi;
const RANGE_SEPARATOR_PATTERN = /\s+(?:to|until|through|-)\s+/i;

type CalDavContext = IExecuteFunctions | ILoadOptionsFunctions;
type CalDavMethod = IHttpRequestMethods | 'PROPFIND' | 'REPORT' | 'MKCALENDAR';
type ResourceLocatorValue = string | { mode?: string; value?: string };

interface CalDavCredentials extends IDataObject {
	baseUrl: string;
	calendarHomePath?: string;
	defaultTimezone?: string;
	ignoreTlsErrors?: boolean;
}

interface XmlResponse {
	href: string;
	statuses: string[];
	properties: IDataObject;
}

interface CalendarInfo extends IDataObject {
	href: string;
	displayName: string;
	name: string;
	description?: string;
	resourceTypes: string[];
	supportedComponents: string[];
	etag?: string;
	ctag?: string;
	rawProperties: IDataObject;
}

interface ParsedCalendarProperty {
	name: string;
	params: Record<string, string>;
	value: string;
}

interface ParsedCalendarComponent {
	name: string;
	properties: ParsedCalendarProperty[];
	children: ParsedCalendarComponent[];
}

interface ParsedCalendarData {
	root: ParsedCalendarComponent;
	events: EventInfo[];
}

export interface EventInfo extends IDataObject {
	href: string;
	uid: string;
	summary: string;
	description?: string;
	location?: string;
	status?: string;
	transparency?: string;
	start: string;
	end?: string;
	duration?: string;
	timezone?: string;
	allDay: boolean;
	sequence?: number;
	organizer?: string;
	attendees: string[];
	rawICalendar: string;
	etag?: string;
}

interface TimeWindow {
	start: DateTime;
	end: DateTime;
	allDay: boolean;
	zone: string;
}

interface StructuredEventInput {
	summary: string;
	description?: string;
	location?: string;
	status?: string;
	transparency?: string;
	start: DateTime;
	end?: DateTime;
	duration?: Duration;
	allDay: boolean;
	timezone: string;
	uid: string;
	attendees: string[];
}

interface EventQueryOptions {
	start?: DateTime;
	end?: DateTime;
	expand?: boolean;
	includeRawICalendar?: boolean;
	limit?: number;
	filters?: {
		uid?: string;
		summaryContains?: string;
		descriptionContains?: string;
		locationContains?: string;
		status?: string;
		attendeeEmail?: string;
	};
}

interface FreeBusyOptions {
	includeTentativeAsBusy: boolean;
	minimumDurationMinutes: number;
	slotMinutes?: number;
}

interface BusyInterval {
	start: string;
	end: string;
}

interface FreeSlot {
	start: string;
	end: string;
	durationMinutes: number;
}

interface FreeBusyResult extends IDataObject {
	windowStart: string;
	windowEnd: string;
	busy: BusyInterval[];
	free: FreeSlot[];
	canFitRequestedDuration: boolean;
	humanSummary: string;
}

function isObject(value: unknown): value is Record<string, unknown> {
	return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function toArray<T>(value: T | T[] | undefined | null): T[] {
	if (value === undefined || value === null) {
		return [];
	}

	return Array.isArray(value) ? value : [value];
}

function escapeXml(value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

function escapeIcalText(value: string): string {
	return value
		.replace(/\\/g, '\\\\')
		.replace(/\n/g, '\\n')
		.replace(/;/g, '\\;')
		.replace(/,/g, '\\,');
}

function unescapeIcalText(value: string): string {
	return value
		.replace(/\\n/g, '\n')
		.replace(/\\N/g, '\n')
		.replace(/\\,/g, ',')
		.replace(/\\;/g, ';')
		.replace(/\\\\/g, '\\');
}

function foldIcalLine(line: string): string {
	const limit = 75;

	if (line.length <= limit) {
		return line;
	}

	const chunks: string[] = [];
	let remaining = line;

	while (remaining.length > limit) {
		chunks.push(remaining.slice(0, limit));
		remaining = ` ${remaining.slice(limit)}`;
	}

	chunks.push(remaining);

	return chunks.join('\r\n');
}

function formatIcalDateTime(dateTime: DateTime, allDay: boolean, timezone: string): string {
	if (allDay) {
		return `;VALUE=DATE:${dateTime.setZone(timezone).toFormat('yyyyLLdd')}`;
	}

	if (timezone !== DEFAULT_TIMEZONE) {
		return `;TZID=${timezone}:${dateTime.setZone(timezone).toFormat("yyyyLLdd'T'HHmmss")}`;
	}

	return `:${dateTime.toUTC().toFormat("yyyyLLdd'T'HHmmss'Z'")}`;
}

function extractText(value: unknown): string {
	if (value === undefined || value === null) {
		return '';
	}

	if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
		return String(value);
	}

	if (Array.isArray(value)) {
		return value.map((entry) => extractText(entry)).find((entry) => entry.length > 0) ?? '';
	}

	if (isObject(value)) {
		if (typeof value._ === 'string') {
			return value._;
		}

		if (typeof value.href === 'string') {
			return value.href;
		}

		for (const entry of Object.values(value)) {
			const text = extractText(entry);
			if (text.length > 0) {
				return text;
			}
		}
	}

	return '';
}

function decodeXmlEntities(value: string): string {
	return value
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&');
}

function escapeRegex(value: string): string {
	return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getTagBlocks(xml: string, localName: string): string[] {
	const expression = new RegExp(
		`<(?:[\\w.-]+:)?${escapeRegex(localName)}(?:\\s[^>]*)?>([\\s\\S]*?)<\\/(?:[\\w.-]+:)?${escapeRegex(localName)}>`,
		'gi',
	);
	const matches: string[] = [];
	let match: RegExpExecArray | null;

	match = expression.exec(xml);
	while (match) {
		matches.push(match[1]);
		match = expression.exec(xml);
	}

	return matches;
}

function getFirstTagBlock(xml: string, localName: string): string {
	return getTagBlocks(xml, localName)[0] ?? '';
}

function stripTags(xml: string): string {
	return decodeXmlEntities(xml.replace(/<[^>]+>/g, '').trim());
}

function getFirstTagText(xml: string, localName: string): string {
	return stripTags(getFirstTagBlock(xml, localName));
}

function getTagAttributeValues(xml: string, localName: string, attributeName: string): string[] {
	const expression = new RegExp(
		`<(?:[\\w.-]+:)?${escapeRegex(localName)}\\b[^>]*\\b${escapeRegex(attributeName)}="([^"]+)"[^>]*\\/?>`,
		'gi',
	);
	const matches: string[] = [];
	let match: RegExpExecArray | null;

	match = expression.exec(xml);
	while (match) {
		matches.push(decodeXmlEntities(match[1]));
		match = expression.exec(xml);
	}

	return matches;
}

function getChildTagNames(xml: string): string[] {
	const expression = /<(?:[\w.-]+:)?([\w.-]+)\b[^>]*\/?>/gi;
	const names = new Set<string>();
	let match: RegExpExecArray | null;

	match = expression.exec(xml);
	while (match) {
		names.add(match[1]);
		match = expression.exec(xml);
	}

	return [...names];
}

function parseResponseProperties(propXml: string): IDataObject {
	const displayName = getFirstTagText(propXml, 'displayname');
	const calendarDescription = getFirstTagText(propXml, 'calendar-description');
	const currentUserPrincipal = getFirstTagText(getFirstTagBlock(propXml, 'current-user-principal'), 'href');
	const calendarHomeSet = getFirstTagText(getFirstTagBlock(propXml, 'calendar-home-set'), 'href');
	const calendarData = decodeXmlEntities(getFirstTagBlock(propXml, 'calendar-data')).trim();

	return {
		displayname: displayName,
		'calendar-description': calendarDescription,
		'current-user-principal': currentUserPrincipal,
		'calendar-home-set': calendarHomeSet,
		getetag: getFirstTagText(propXml, 'getetag'),
		getctag: getFirstTagText(propXml, 'getctag'),
		'calendar-data': calendarData,
		resourcetype: getChildTagNames(getFirstTagBlock(propXml, 'resourcetype')),
		'supported-calendar-component-set': getTagAttributeValues(
			getFirstTagBlock(propXml, 'supported-calendar-component-set'),
			'comp',
			'name',
		),
	};
}

function parseXmlResponses(xml: string): XmlResponse[] {
	return getTagBlocks(xml, 'response').map((responseXml) => {
		const propstats = getTagBlocks(responseXml, 'propstat');
		const properties: IDataObject = {};
		const statuses: string[] = [];

		for (const propstatXml of propstats) {
			const status = getFirstTagText(propstatXml, 'status');
			statuses.push(status);

			if (!status.includes(' 200 ')) {
				continue;
			}

			const parsedProperties = parseResponseProperties(getFirstTagBlock(propstatXml, 'prop'));
			Object.assign(properties, parsedProperties);
		}

		return {
			href: getFirstTagText(responseXml, 'href'),
			statuses,
			properties,
		};
	});
}

function getCredentialsTimezone(credentials: CalDavCredentials): string {
	return typeof credentials.defaultTimezone === 'string' && credentials.defaultTimezone.trim().length > 0
		? credentials.defaultTimezone.trim()
		: DEFAULT_TIMEZONE;
}

function isAbsoluteUrl(value: string): boolean {
	return /^https?:\/\//i.test(value);
}

function ensureTrailingSlash(value: string): string {
	return value.endsWith('/') ? value : `${value}/`;
}

function normalizeBaseUrl(value: string): string {
	const trimmed = value.trim();
	return trimmed.endsWith('/') ? trimmed.slice(0, -1) : trimmed;
}

function toAbsoluteUrl(baseUrl: string, value: string, collection = false): string {
	const normalizedBaseUrl = ensureTrailingSlash(normalizeBaseUrl(baseUrl));
	const resolved = new URL(value, normalizedBaseUrl).toString();
	return collection ? ensureTrailingSlash(resolved) : resolved;
}

function toCollectionUrl(baseUrl: string, value: string): string {
	return toAbsoluteUrl(baseUrl, value, true);
}

function getStringLocatorValue(value: ResourceLocatorValue): string {
	if (typeof value === 'string') {
		return value.trim();
	}

	if (isObject(value) && typeof value.value === 'string') {
		return value.value.trim();
	}

	return '';
}

function getXmlPropertyTag(name: string): string {
	if (name === 'getctag') {
		return 'CS:getctag';
	}

	return `${DAV_PROPERTY_NAMES.has(name) ? 'D' : 'C'}:${name}`;
}

function buildPropfindBody(propertyNames: string[]): string {
	const tags = propertyNames.map((name) => `<${getXmlPropertyTag(name)} />`).join('');

	return [
		'<?xml version="1.0" encoding="UTF-8"?>',
		`<D:propfind xmlns:D="${DAV_NAMESPACE}" xmlns:C="${CALDAV_NAMESPACE}" xmlns:CS="${CALENDAR_SERVER_NAMESPACE}">`,
		'<D:prop>',
		tags,
		'</D:prop>',
		'</D:propfind>',
	].join('');
}

async function calDavRequest<T = unknown>(
	context: CalDavContext,
	requestOptions: IHttpRequestOptions,
	method: CalDavMethod,
): Promise<T> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;

	const preparedOptions: IHttpRequestOptions = {
		...requestOptions,
		method: method as unknown as IHttpRequestMethods,
		headers: {
			...requestOptions.headers,
		},
	};

	if (credentials.ignoreTlsErrors === true) {
		preparedOptions.skipSslCertificateValidation = true;
	}

	return (await context.helpers.httpRequestWithAuthentication.call(
		context,
		'calDavApi',
		preparedOptions,
	)) as T;
}

async function propfind(
	context: CalDavContext,
	url: string,
	depth: '0' | '1',
	propertyNames: string[],
): Promise<XmlResponse[]> {
	const body = buildPropfindBody(propertyNames);
	const response = await calDavRequest<string>(
		context,
		{
			url,
			body,
			headers: {
				Depth: depth,
				'Content-Type': 'application/xml; charset=utf-8',
			},
			returnFullResponse: false,
			json: false,
		},
		'PROPFIND',
	);

	return parseXmlResponses(response);
}

function extractResourceTypes(value: unknown): string[] {
	if (Array.isArray(value)) {
		return value.filter((entry): entry is string => typeof entry === 'string');
	}

	if (!isObject(value)) {
		return [];
	}

	return Object.keys(value).filter((key) => key !== '_');
}

function extractSupportedComponents(value: unknown): string[] {
	if (Array.isArray(value)) {
		return value.filter((entry): entry is string => typeof entry === 'string');
	}

	if (!isObject(value)) {
		return [];
	}

	const components = toArray(value.comp).filter((entry): entry is Record<string, unknown> => isObject(entry));

	return components
		.map((entry) => (typeof entry.name === 'string' ? entry.name : ''))
		.filter((entry) => entry.length > 0);
}

async function discoverCalendarHome(context: CalDavContext): Promise<string> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;
	const baseUrl = normalizeBaseUrl(credentials.baseUrl);

	if (typeof credentials.calendarHomePath === 'string' && credentials.calendarHomePath.trim().length > 0) {
		return toCollectionUrl(baseUrl, credentials.calendarHomePath.trim());
	}

	const baseResponses = await propfind(context, ensureTrailingSlash(baseUrl), '0', [
		'current-user-principal',
		'calendar-home-set',
		'resourcetype',
	]);
	const baseResponse = baseResponses[0];

	const homeSet = extractText(baseResponse?.properties['calendar-home-set']);
	if (homeSet.length > 0) {
		return toCollectionUrl(baseUrl, homeSet);
	}

	const principalHref = extractText(baseResponse?.properties['current-user-principal']);
	if (principalHref.length > 0) {
		const principalResponses = await propfind(context, toAbsoluteUrl(baseUrl, principalHref), '0', [
			'calendar-home-set',
		]);
		const discoveredHome = extractText(principalResponses[0]?.properties['calendar-home-set']);

		if (discoveredHome.length > 0) {
			return toCollectionUrl(baseUrl, discoveredHome);
		}
	}

	return ensureTrailingSlash(baseUrl);
}

function mapCalendarInfo(baseUrl: string, response: XmlResponse): CalendarInfo {
	const displayName = extractText(response.properties.displayname);
	const href = toCollectionUrl(baseUrl, response.href);

	return {
		href,
		displayName: displayName || href,
		name: href.replace(/\/$/, '').split('/').pop() ?? href,
		description: extractText(response.properties['calendar-description']) || undefined,
		resourceTypes: extractResourceTypes(response.properties.resourcetype),
		supportedComponents: extractSupportedComponents(response.properties['supported-calendar-component-set']),
		etag: extractText(response.properties.getetag) || undefined,
		ctag: extractText(response.properties.getctag) || undefined,
		rawProperties: response.properties,
	};
}

export async function listCalendars(context: CalDavContext): Promise<CalendarInfo[]> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;
	const baseUrl = normalizeBaseUrl(credentials.baseUrl);
	const homeUrl = await discoverCalendarHome(context);
	const responses = await propfind(context, homeUrl, '1', [
		'displayname',
		'resourcetype',
		'calendar-description',
		'supported-calendar-component-set',
		'getetag',
		'getctag',
	]);

	const calendars = responses
		.map((response) => mapCalendarInfo(baseUrl, response))
		.filter((calendar) => calendar.resourceTypes.includes('calendar'));

	if (calendars.length > 0) {
		return deduplicateCalendars(calendars);
	}

	const rootResponse = responses.find((response) => toCollectionUrl(baseUrl, response.href) === homeUrl);
	if (!rootResponse) {
		return [];
	}

	const rootCalendar = mapCalendarInfo(baseUrl, rootResponse);
	return rootCalendar.resourceTypes.includes('calendar') ? [rootCalendar] : [];
}

function deduplicateCalendars(calendars: CalendarInfo[]): CalendarInfo[] {
	const seen = new Set<string>();
	return calendars.filter((calendar) => {
		if (seen.has(calendar.href)) {
			return false;
		}

		seen.add(calendar.href);
		return true;
	});
}

export async function searchCalendars(
	context: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const normalizedFilter = filter?.trim().toLowerCase() ?? '';
	const calendars = await listCalendars(context);

	return {
		results: calendars
			.filter((calendar) => {
				if (normalizedFilter.length === 0) {
					return true;
				}

				return [calendar.displayName, calendar.name, calendar.href]
					.map((entry) => entry.toLowerCase())
					.some((entry) => entry.includes(normalizedFilter));
			})
			.map((calendar) => ({
				name: calendar.displayName,
				value: calendar.href,
				description: calendar.description ?? calendar.href,
				url: calendar.href,
			})),
	};
}

export async function resolveCalendarHref(
	context: CalDavContext,
	locatorValue: ResourceLocatorValue,
): Promise<string> {
	const rawValue = getStringLocatorValue(locatorValue);
	if (rawValue.length === 0) {
		throw new Error('A calendar is required for this operation.');
	}

	if (isAbsoluteUrl(rawValue)) {
		return ensureTrailingSlash(rawValue);
	}

	const homeUrl = await discoverCalendarHome(context);
	return toCollectionUrl(homeUrl, rawValue);
}

export async function getCalendar(context: CalDavContext, locatorValue: ResourceLocatorValue): Promise<CalendarInfo> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;
	const baseUrl = normalizeBaseUrl(credentials.baseUrl);
	const calendarHref = await resolveCalendarHref(context, locatorValue);
	const responses = await propfind(context, calendarHref, '0', [
		'displayname',
		'resourcetype',
		'calendar-description',
		'supported-calendar-component-set',
		'getetag',
		'getctag',
	]);
	const response = responses[0];

	if (!response) {
		throw new Error(`Unable to read calendar properties for ${calendarHref}.`);
	}

	return mapCalendarInfo(baseUrl, response);
}

export async function createCalendar(
	context: CalDavContext,
	options: {
		path: string;
		displayName: string;
		description?: string;
	},
): Promise<CalendarInfo> {
	const homeUrl = await discoverCalendarHome(context);
	const normalizedPath = options.path.trim().replace(/^\/+/, '').replace(/\/+$/, '');

	if (normalizedPath.length === 0) {
		throw new Error('Calendar Path cannot be empty.');
	}

	const targetUrl = ensureTrailingSlash(new URL(`${normalizedPath}/`, homeUrl).toString());
	const body = [
		'<?xml version="1.0" encoding="UTF-8"?>',
		`<C:mkcalendar xmlns:D="${DAV_NAMESPACE}" xmlns:C="${CALDAV_NAMESPACE}">`,
		'<D:set>',
		'<D:prop>',
		`<D:displayname>${escapeXml(options.displayName)}</D:displayname>`,
		...(options.description
			? [`<C:calendar-description>${escapeXml(options.description)}</C:calendar-description>`]
			: []),
		'<C:supported-calendar-component-set>',
		'<C:comp name="VEVENT" />',
		'</C:supported-calendar-component-set>',
		'</D:prop>',
		'</D:set>',
		'</C:mkcalendar>',
	].join('');

	await calDavRequest(context, {
		url: targetUrl,
		body,
		headers: {
			'Content-Type': 'application/xml; charset=utf-8',
		},
		json: false,
	}, 'MKCALENDAR');

	return await getCalendar(context, targetUrl);
}

export async function deleteCalendar(context: CalDavContext, locatorValue: ResourceLocatorValue): Promise<void> {
	const calendarHref = await resolveCalendarHref(context, locatorValue);

	await calDavRequest(context, {
		url: calendarHref,
	}, 'DELETE');
}

function unfoldIcal(rawICalendar: string): string[] {
	return rawICalendar
		.replace(/\r\n[ \t]/g, '')
		.replace(/\n[ \t]/g, '')
		.split(/\r?\n/)
		.map((line) => line.trimEnd())
		.filter((line) => line.length > 0);
}

function parsePropertyLine(line: string): ParsedCalendarProperty {
	const separatorIndex = line.indexOf(':');

	if (separatorIndex === -1) {
		return {
			name: line.toUpperCase(),
			params: {},
			value: '',
		};
	}

	const descriptor = line.slice(0, separatorIndex);
	const value = line.slice(separatorIndex + 1);
	const segments = descriptor.split(';');
	const name = segments.shift()?.toUpperCase() ?? descriptor.toUpperCase();
	const params: Record<string, string> = {};

	for (const segment of segments) {
		const [key, ...rawValueParts] = segment.split('=');
		params[key.toUpperCase()] = rawValueParts.join('=');
	}

	return {
		name,
		params,
		value,
	};
}

function parseICalendar(rawICalendar: string, defaultTimezone: string, href = '', etag?: string): ParsedCalendarData {
	const lines = unfoldIcal(rawICalendar);
	const root: ParsedCalendarComponent = {
		name: 'ROOT',
		properties: [],
		children: [],
	};
	const stack: ParsedCalendarComponent[] = [root];

	for (const line of lines) {
		if (line.startsWith('BEGIN:')) {
			const component: ParsedCalendarComponent = {
				name: line.slice('BEGIN:'.length).toUpperCase(),
				properties: [],
				children: [],
			};

			stack[stack.length - 1].children.push(component);
			stack.push(component);
			continue;
		}

		if (line.startsWith('END:')) {
			stack.pop();
			continue;
		}

		stack[stack.length - 1].properties.push(parsePropertyLine(line));
	}

	const events = collectComponents(root, 'VEVENT').map((eventComponent) =>
		mapEventComponent(eventComponent, rawICalendar, defaultTimezone, href, etag),
	);

	return {
		root,
		events,
	};
}

function collectComponents(component: ParsedCalendarComponent, name: string): ParsedCalendarComponent[] {
	const matches: ParsedCalendarComponent[] = [];

	for (const child of component.children) {
		if (child.name === name) {
			matches.push(child);
		}

		matches.push(...collectComponents(child, name));
	}

	return matches;
}

function getProperty(component: ParsedCalendarComponent, name: string): ParsedCalendarProperty | undefined {
	return component.properties.find((property) => property.name === name);
}

function getProperties(component: ParsedCalendarComponent, name: string): ParsedCalendarProperty[] {
	return component.properties.filter((property) => property.name === name);
}

function parseIcalDateTime(
	property: ParsedCalendarProperty | undefined,
	defaultTimezone: string,
): { value?: DateTime; allDay: boolean; timezone: string } {
	if (!property) {
		return { allDay: false, timezone: defaultTimezone };
	}

	const timezone = property.params.TZID ?? defaultTimezone;
	const value = property.value.trim();

	if (property.params.VALUE === 'DATE' || /^\d{8}$/.test(value)) {
		return {
			value: DateTime.fromFormat(value, 'yyyyLLdd', { zone: timezone }).startOf('day'),
			allDay: true,
			timezone,
		};
	}

	const formats = ["yyyyLLdd'T'HHmmss'Z'", "yyyyLLdd'T'HHmm'Z'", "yyyyLLdd'T'HHmmss", "yyyyLLdd'T'HHmm"];

	for (const format of formats) {
		const candidate = DateTime.fromFormat(value, format, {
			zone: value.endsWith('Z') ? DEFAULT_TIMEZONE : timezone,
		});

		if (candidate.isValid) {
			return {
				value: value.endsWith('Z') ? candidate.toUTC() : candidate.setZone(timezone),
				allDay: false,
				timezone,
			};
		}
	}

	return { allDay: false, timezone };
}

function mapEventComponent(
	component: ParsedCalendarComponent,
	rawICalendar: string,
	defaultTimezone: string,
	href: string,
	etag?: string,
): EventInfo {
	const startProperty = getProperty(component, 'DTSTART');
	const endProperty = getProperty(component, 'DTEND');
	const durationProperty = getProperty(component, 'DURATION');
	const start = parseIcalDateTime(startProperty, defaultTimezone);
	const end = parseIcalDateTime(endProperty, start.timezone);
	const attendees = getProperties(component, 'ATTENDEE')
		.map((property) => property.value.replace(/^mailto:/i, ''))
		.filter((value) => value.length > 0);
	const sequence = Number.parseInt(getProperty(component, 'SEQUENCE')?.value ?? '', 10);
	const organizer = getProperty(component, 'ORGANIZER')?.value.replace(/^mailto:/i, '');

	return {
		href,
		uid: getProperty(component, 'UID')?.value ?? randomUUID(),
		summary: unescapeIcalText(getProperty(component, 'SUMMARY')?.value ?? ''),
		description: getProperty(component, 'DESCRIPTION')
			? unescapeIcalText(getProperty(component, 'DESCRIPTION')?.value ?? '')
			: undefined,
		location: getProperty(component, 'LOCATION')
			? unescapeIcalText(getProperty(component, 'LOCATION')?.value ?? '')
			: undefined,
		status: getProperty(component, 'STATUS')?.value,
		transparency: getProperty(component, 'TRANSP')?.value,
		start: start.value?.toISO() ?? '',
		end: end.value?.toISO() ?? undefined,
		duration: durationProperty?.value,
		timezone: start.timezone,
		allDay: start.allDay,
		sequence: Number.isFinite(sequence) ? sequence : undefined,
		organizer: organizer?.length ? organizer : undefined,
		attendees,
		rawICalendar,
		etag,
	};
}

function serializeCalendarProperty(
	name: string,
	value: string,
	params: Record<string, string> = {},
): string {
	const paramString = Object.entries(params)
		.map(([key, nestedValue]) => `;${key}=${nestedValue}`)
		.join('');

	return foldIcalLine(`${name}${paramString}:${value}`);
}

function buildEventCalendar(input: StructuredEventInput): string {
	const lines = [
		'BEGIN:VCALENDAR',
		'VERSION:2.0',
		`PRODID:${PROD_ID}`,
		'CALSCALE:GREGORIAN',
		'BEGIN:VEVENT',
		serializeCalendarProperty('UID', input.uid),
		serializeCalendarProperty(
			'DTSTAMP',
			DateTime.utc().toFormat("yyyyLLdd'T'HHmmss'Z'"),
		),
		foldIcalLine(`DTSTART${formatIcalDateTime(input.start, input.allDay, input.timezone)}`),
	];

	if (input.end) {
		lines.push(foldIcalLine(`DTEND${formatIcalDateTime(input.end, input.allDay, input.timezone)}`));
	} else if (input.duration) {
		lines.push(serializeCalendarProperty('DURATION', input.duration.toISO()));
	}

	if (input.summary.length > 0) {
		lines.push(serializeCalendarProperty('SUMMARY', escapeIcalText(input.summary)));
	}

	if (input.description) {
		lines.push(serializeCalendarProperty('DESCRIPTION', escapeIcalText(input.description)));
	}

	if (input.location) {
		lines.push(serializeCalendarProperty('LOCATION', escapeIcalText(input.location)));
	}

	if (input.status) {
		lines.push(serializeCalendarProperty('STATUS', input.status.toUpperCase()));
	}

	if (input.transparency) {
		lines.push(serializeCalendarProperty('TRANSP', input.transparency.toUpperCase()));
	}

	for (const attendee of input.attendees) {
		lines.push(serializeCalendarProperty('ATTENDEE', `mailto:${attendee}`));
	}

	lines.push('END:VEVENT', 'END:VCALENDAR');

	return `${lines.join('\r\n')}\r\n`;
}

function replaceEventProperties(
	existingICalendar: string,
	properties: Map<string, string[]>,
): string {
	const lines = unfoldIcal(existingICalendar);
	const updatedLines: string[] = [];
	let insideFirstEvent = false;
	let eventHandled = false;

	for (const line of lines) {
		if (line === 'BEGIN:VEVENT') {
			insideFirstEvent = true;
			eventHandled = true;
			updatedLines.push(line);
			continue;
		}

		if (line === 'END:VEVENT' && insideFirstEvent) {
			for (const values of properties.values()) {
				updatedLines.push(...values);
			}

			insideFirstEvent = false;
			updatedLines.push(line);
			continue;
		}

		if (insideFirstEvent) {
			const propertyName = line.split(':')[0].split(';')[0].toUpperCase();
			if (properties.has(propertyName)) {
				continue;
			}
		}

		updatedLines.push(line);
	}

	if (!eventHandled) {
		throw new Error('The existing iCalendar payload does not contain a VEVENT component.');
	}

	return `${updatedLines.join('\r\n')}\r\n`;
}

function getFilenameForEvent(uid: string, explicitFilename?: string): string {
	const value = explicitFilename?.trim() || `${uid}.ics`;
	return value.endsWith('.ics') ? value : `${value}.ics`;
}

function parseTimeToken(value: string): { hour: number; minute: number } | undefined {
	const normalizedValue = value.trim().toLowerCase();

	if (normalizedValue === 'noon') {
		return { hour: 12, minute: 0 };
	}

	if (normalizedValue === 'midnight') {
		return { hour: 0, minute: 0 };
	}

	const match = normalizedValue.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
	if (!match) {
		return undefined;
	}

	let hour = Number.parseInt(match[1], 10);
	const minute = Number.parseInt(match[2] ?? '0', 10);
	const meridiem = match[3]?.toLowerCase();

	if (meridiem === 'pm' && hour < 12) {
		hour += 12;
	}

	if (meridiem === 'am' && hour === 12) {
		hour = 0;
	}

	if (hour > 23 || minute > 59) {
		return undefined;
	}

	return { hour, minute };
}

function stripOrdinals(value: string): string {
	return value.replace(ORDINAL_SUFFIX_PATTERN, '$1');
}

function applyTimeToDate(date: DateTime, time: { hour: number; minute: number }): DateTime {
	return date.set({
		hour: time.hour,
		minute: time.minute,
		second: 0,
		millisecond: 0,
	});
}

function resolveWeekday(baseDate: DateTime, weekday: string, modifier: 'next' | 'this' | 'last' | undefined): DateTime {
	const targetWeekday = DAYS_OF_WEEK.indexOf(weekday as (typeof DAYS_OF_WEEK)[number]) + 1;
	const candidate = baseDate.startOf('day');
	let delta = targetWeekday - candidate.weekday;

	if (modifier === 'next' && delta <= 0) {
		delta += 7;
	}

	if (modifier === 'last' && delta >= 0) {
		delta -= 7;
	}

	if (!modifier && delta < 0) {
		delta += 7;
	}

	return candidate.plus({ days: delta });
}

function parseFlexibleDateTime(
	value: string,
	timezone: string,
	baseDate: DateTime,
	referenceDate?: DateTime,
): DateTime | undefined {
	const trimmedValue = value.trim();
	if (trimmedValue.length === 0) {
		return undefined;
	}

	const dateTimeFromIso = DateTime.fromISO(trimmedValue, { zone: timezone });
	if (dateTimeFromIso.isValid) {
		return dateTimeFromIso;
	}

	const jsDate = new Date(trimmedValue);
	if (!Number.isNaN(jsDate.valueOf())) {
		return DateTime.fromJSDate(jsDate, { zone: timezone });
	}

	const timeOnly = parseTimeToken(trimmedValue);
	if (timeOnly && referenceDate) {
		return applyTimeToDate(referenceDate.setZone(timezone), timeOnly);
	}

	const normalizedValue = stripOrdinals(trimmedValue.replace(/,/g, '')).toLowerCase();

	for (const format of [
		'yyyy-LL-dd HH:mm',
		'yyyy-LL-dd H:mm',
		'yyyy-LL-dd HH:mm:ss',
		'yyyy-LL-dd',
		'LL/dd/yyyy HH:mm',
		'LL/dd/yyyy',
		'LL/dd/yy HH:mm',
		'LL/dd/yy',
		'LLLL d yyyy h:mm a',
		'LLLL d yyyy',
		'LLLL d h:mm a',
		'LLLL d',
		'LLL d yyyy h:mm a',
		'LLL d yyyy',
		'LLL d h:mm a',
		'LLL d',
	]) {
		const candidate = DateTime.fromFormat(trimmedValue, format, { zone: timezone });
		if (candidate.isValid) {
			if (!format.includes('yyyy')) {
				const withYear = candidate.set({ year: baseDate.year });
				return withYear < baseDate.startOf('day') ? withYear.plus({ years: 1 }) : withYear;
			}

			return candidate;
		}
	}

	const relativeMatch = normalizedValue.match(/^(today|tomorrow|yesterday)(?:\s+at\s+(.+))?$/i);
	if (relativeMatch) {
		const anchor = baseDate.startOf('day').plus({
			days:
				relativeMatch[1] === 'tomorrow' ? 1 : relativeMatch[1] === 'yesterday' ? -1 : 0,
		});
		const parsedTime = relativeMatch[2] ? parseTimeToken(relativeMatch[2]) : undefined;
		return parsedTime ? applyTimeToDate(anchor, parsedTime) : anchor;
	}

	const weekdayMatch = normalizedValue.match(
		/^(?:(next|this|last)\s+)?(monday|tuesday|wednesday|thursday|friday|saturday|sunday)(?:\s+at\s+(.+))?$/,
	);
	if (weekdayMatch) {
		const date = resolveWeekday(
			baseDate,
			weekdayMatch[2],
			weekdayMatch[1] as 'next' | 'this' | 'last' | undefined,
		);
		const parsedTime = weekdayMatch[3] ? parseTimeToken(weekdayMatch[3]) : undefined;
		return parsedTime ? applyTimeToDate(date, parsedTime) : date;
	}

	return undefined;
}

function parseFlexibleDuration(value: string): Duration | undefined {
	const trimmedValue = value.trim();
	if (trimmedValue.length === 0) {
		return undefined;
	}

	const isoDuration = Duration.fromISO(trimmedValue);
	if (isoDuration.isValid) {
		return isoDuration;
	}

	const durationMatch = trimmedValue.match(
		/^(?:(\d+)\s*h(?:ours?)?)?\s*(?:(\d+)\s*m(?:in(?:utes?)?)?)?$/i,
	);

	if (!durationMatch) {
		return undefined;
	}

	const hours = Number.parseInt(durationMatch[1] ?? '0', 10);
	const minutes = Number.parseInt(durationMatch[2] ?? '0', 10);
	if (hours === 0 && minutes === 0) {
		return undefined;
	}

	return Duration.fromObject({ hours, minutes });
}

function getPartOfDayWindow(partOfDay: string): { startHour: number; startMinute: number; endHour: number; endMinute: number } | undefined {
	switch (partOfDay.toLowerCase()) {
		case 'morning':
			return { startHour: 9, startMinute: 0, endHour: 12, endMinute: 0 };
		case 'afternoon':
			return { startHour: 13, startMinute: 0, endHour: 17, endMinute: 0 };
		case 'evening':
			return { startHour: 17, startMinute: 0, endHour: 21, endMinute: 0 };
		case 'night':
			return { startHour: 20, startMinute: 0, endHour: 23, endMinute: 0 };
		default:
			return undefined;
	}
}

export function resolveTimeWindow(
	rangeText: string | undefined,
	startText: string | undefined,
	endText: string | undefined,
	timezone: string,
): TimeWindow {
	const baseDate = DateTime.now().setZone(timezone);

	if (startText?.trim()) {
		const start = parseFlexibleDateTime(startText, timezone, baseDate);
		if (!start) {
			throw new Error(`Unable to parse start value "${startText}".`);
		}

		let end: DateTime | undefined;
		if (endText?.trim()) {
			end = parseFlexibleDateTime(endText, timezone, baseDate, start);
			if (!end) {
				throw new Error(`Unable to parse end value "${endText}".`);
			}
		}

		return {
			start,
			end: end ?? start.plus({ minutes: DEFAULT_EVENT_DURATION_MINUTES }),
			allDay: false,
			zone: timezone,
		};
	}

	if (!rangeText?.trim()) {
		throw new Error('A date or time range is required.');
	}

	const normalizedRangeText = rangeText.trim().toLowerCase();

	if (normalizedRangeText === 'today') {
		return {
			start: baseDate.startOf('day'),
			end: baseDate.endOf('day'),
			allDay: true,
			zone: timezone,
		};
	}

	if (normalizedRangeText === 'tomorrow') {
		const tomorrow = baseDate.plus({ days: 1 });
		return {
			start: tomorrow.startOf('day'),
			end: tomorrow.endOf('day'),
			allDay: true,
			zone: timezone,
		};
	}

	if (normalizedRangeText === 'this week' || normalizedRangeText === 'next week') {
		const offset = normalizedRangeText === 'next week' ? 1 : 0;
		const weekStart = baseDate.plus({ weeks: offset }).startOf('week');
		return {
			start: weekStart,
			end: weekStart.endOf('week'),
			allDay: true,
			zone: timezone,
		};
	}

	const partOfDayMatch = normalizedRangeText.match(
		/^(?:(next|this|last)\s+)?(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s+(morning|afternoon|evening|night)$/,
	);
	if (partOfDayMatch) {
		const date = resolveWeekday(
			baseDate,
			partOfDayMatch[2],
			partOfDayMatch[1] as 'next' | 'this' | 'last' | undefined,
		);
		const window = getPartOfDayWindow(partOfDayMatch[3]);
		if (!window) {
			throw new Error(`Unsupported part of day "${partOfDayMatch[3]}".`);
		}

		return {
			start: date.set({
				hour: window.startHour,
				minute: window.startMinute,
				second: 0,
				millisecond: 0,
			}),
			end: date.set({
				hour: window.endHour,
				minute: window.endMinute,
				second: 0,
				millisecond: 0,
			}),
			allDay: false,
			zone: timezone,
		};
	}

	const rangeParts = rangeText.split(RANGE_SEPARATOR_PATTERN);
	if (rangeParts.length === 2) {
		const start = parseFlexibleDateTime(rangeParts[0], timezone, baseDate);
		if (!start) {
			throw new Error(`Unable to parse start value "${rangeParts[0]}".`);
		}

		const end = parseFlexibleDateTime(rangeParts[1], timezone, baseDate, start);
		if (!end) {
			throw new Error(`Unable to parse end value "${rangeParts[1]}".`);
		}

		return {
			start,
			end,
			allDay: false,
			zone: timezone,
		};
	}

	const dateOnly = parseFlexibleDateTime(rangeText, timezone, baseDate);
	if (dateOnly) {
		return {
			start: dateOnly.startOf('day'),
			end: dateOnly.endOf('day'),
			allDay: true,
			zone: timezone,
		};
	}

	throw new Error(`Unable to parse time range "${rangeText}".`);
}

export function resolveEventInput(
	options: {
		title: string;
		whenText?: string;
		start?: string;
		end?: string;
		duration?: string;
		allDay?: boolean;
		timezone: string;
		description?: string;
		location?: string;
		status?: string;
		transparency?: string;
		uid?: string;
		attendeesText?: string;
		existingEvent?: EventInfo;
	},
): StructuredEventInput {
	const timezone = options.timezone || DEFAULT_TIMEZONE;
	const baseDate = DateTime.now().setZone(timezone);
	const allDay = options.allDay === true;
	let start: DateTime | undefined;
	let end: DateTime | undefined;

	if (options.whenText?.trim()) {
		const rangeText = options.whenText.trim();
		const rangeParts = rangeText.split(RANGE_SEPARATOR_PATTERN);
		const durationMatch = rangeText.match(/^(.*?)\s+for\s+(.+)$/i);

		if (rangeParts.length === 2) {
			start = parseFlexibleDateTime(rangeParts[0], timezone, baseDate);
			if (!start) {
				throw new Error(`Unable to parse the start value in "${options.whenText}".`);
			}

			end = parseFlexibleDateTime(rangeParts[1], timezone, baseDate, start);
			if (!end) {
				throw new Error(`Unable to parse the end value in "${options.whenText}".`);
			}
		} else if (durationMatch) {
			start = parseFlexibleDateTime(durationMatch[1], timezone, baseDate);
			if (!start) {
				throw new Error(`Unable to parse the start value in "${options.whenText}".`);
			}

			const parsedDuration = parseFlexibleDuration(durationMatch[2]);
			if (!parsedDuration) {
				throw new Error(`Unable to parse the duration in "${options.whenText}".`);
			}

			end = start.plus(parsedDuration);
		} else {
			start = parseFlexibleDateTime(rangeText, timezone, baseDate);
			if (!start) {
				throw new Error(`Unable to parse "${options.whenText}".`);
			}
		}
	}

	if (options.start?.trim()) {
		start = parseFlexibleDateTime(options.start, timezone, baseDate);
		if (!start) {
			throw new Error(`Unable to parse start value "${options.start}".`);
		}
	}

	if (!start && options.existingEvent?.start) {
		start = DateTime.fromISO(options.existingEvent.start, { zone: options.existingEvent.timezone ?? timezone });
	}

	if (!start) {
		throw new Error('A start date or natural language schedule is required.');
	}

	if (options.end?.trim()) {
		end = parseFlexibleDateTime(options.end, timezone, baseDate, start);
		if (!end) {
			throw new Error(`Unable to parse end value "${options.end}".`);
		}
	}

	const duration = options.duration?.trim()
		? parseFlexibleDuration(options.duration)
		: undefined;

	if (!end && duration) {
		end = start.plus(duration);
	}

	if (!end && options.existingEvent?.end) {
		const existingEnd = DateTime.fromISO(options.existingEvent.end, {
			zone: options.existingEvent.timezone ?? timezone,
		});
		const existingStart = DateTime.fromISO(options.existingEvent.start, {
			zone: options.existingEvent.timezone ?? timezone,
		});
		if (existingEnd.isValid && existingStart.isValid) {
			end = start.plus(existingEnd.diff(existingStart));
		}
	}

	if (!end) {
		end = allDay ? start.plus({ days: 1 }).startOf('day') : start.plus({ minutes: DEFAULT_EVENT_DURATION_MINUTES });
	}

	const attendees = (options.attendeesText ?? '')
		.split(/[\n,;]/)
		.map((value) => value.trim())
		.filter((value) => value.length > 0);

	return {
		summary: options.title,
		description: options.description || options.existingEvent?.description,
		location: options.location || options.existingEvent?.location,
		status: options.status || options.existingEvent?.status,
		transparency: options.transparency || options.existingEvent?.transparency,
		start,
		end,
		duration: duration && !end ? duration : undefined,
		allDay,
		timezone,
		uid: options.uid?.trim() || options.existingEvent?.uid || randomUUID(),
		attendees: attendees.length > 0 ? attendees : options.existingEvent?.attendees ?? [],
	};
}

export function formatHumanDateTime(value: string, timezone: string, allDay = false): string {
	const dateTime = DateTime.fromISO(value, { zone: timezone });

	if (!dateTime.isValid) {
		return value;
	}

	return allDay
		? dateTime.toFormat('yyyy-LL-dd')
		: dateTime.setZone(timezone).toFormat("yyyy-LL-dd HH:mm ZZZZ");
}

export function describeEvent(event: EventInfo): string {
	const title = event.summary || event.uid;
	const start = formatHumanDateTime(event.start, event.timezone ?? DEFAULT_TIMEZONE, event.allDay);
	const end = event.end
		? formatHumanDateTime(event.end, event.timezone ?? DEFAULT_TIMEZONE, event.allDay)
		: undefined;

	return end ? `"${title}" from ${start} to ${end}` : `"${title}" at ${start}`;
}

function eventMatchesFilters(event: EventInfo, filters: EventQueryOptions['filters']): boolean {
	if (!filters) {
		return true;
	}

	const contains = (haystack: string | undefined, needle: string | undefined) =>
		!needle || (haystack ?? '').toLowerCase().includes(needle.toLowerCase());

	const attendeeMatches =
		!filters.attendeeEmail ||
		event.attendees.some((attendee) => attendee.toLowerCase().includes(filters.attendeeEmail!.toLowerCase()));

	return (
		(!filters.uid || event.uid === filters.uid) &&
		contains(event.summary, filters.summaryContains) &&
		contains(event.description, filters.descriptionContains) &&
		contains(event.location, filters.locationContains) &&
		(!filters.status || (event.status ?? '').toLowerCase() === filters.status.toLowerCase()) &&
		attendeeMatches
	);
}

function buildCalendarQueryReport(options: Pick<EventQueryOptions, 'start' | 'end' | 'expand'>): string {
	const timeRange =
		options.start && options.end
			? `<C:time-range start="${options.start.toUTC().toFormat("yyyyLLdd'T'HHmmss'Z'")}" end="${options.end
					.toUTC()
					.toFormat("yyyyLLdd'T'HHmmss'Z'")}" />`
			: '';

	const calendarData = options.expand && options.start && options.end
		? [
				'<C:calendar-data>',
				`<C:expand start="${options.start.toUTC().toFormat("yyyyLLdd'T'HHmmss'Z'")}" end="${options.end
					.toUTC()
					.toFormat("yyyyLLdd'T'HHmmss'Z'")}" />`,
				'</C:calendar-data>',
			].join('')
		: '<C:calendar-data />';

	return [
		'<?xml version="1.0" encoding="UTF-8"?>',
		`<C:calendar-query xmlns:D="${DAV_NAMESPACE}" xmlns:C="${CALDAV_NAMESPACE}">`,
		'<D:prop>',
		'<D:getetag />',
		calendarData,
		'</D:prop>',
		'<C:filter>',
		'<C:comp-filter name="VCALENDAR">',
		'<C:comp-filter name="VEVENT">',
		timeRange,
		'</C:comp-filter>',
		'</C:comp-filter>',
		'</C:filter>',
		'</C:calendar-query>',
	].join('');
}

function buildUidQueryReport(uid: string): string {
	return [
		'<?xml version="1.0" encoding="UTF-8"?>',
		`<C:calendar-query xmlns:D="${DAV_NAMESPACE}" xmlns:C="${CALDAV_NAMESPACE}">`,
		'<D:prop>',
		'<D:getetag />',
		'<C:calendar-data />',
		'</D:prop>',
		'<C:filter>',
		'<C:comp-filter name="VCALENDAR">',
		'<C:comp-filter name="VEVENT">',
		'<C:prop-filter name="UID">',
		`<C:text-match collation="${TEXT_MATCH_COLLATION}">${escapeXml(uid)}</C:text-match>`,
		'</C:prop-filter>',
		'</C:comp-filter>',
		'</C:comp-filter>',
		'</C:filter>',
		'</C:calendar-query>',
	].join('');
}

function mapEventFromResponse(
	defaultTimezone: string,
	calendarHref: string,
	response: XmlResponse,
): EventInfo[] {
	const rawICalendar = extractText(response.properties['calendar-data']);
	if (!rawICalendar) {
		return [];
	}

	const href = toAbsoluteUrl(calendarHref, response.href);
	const parsed = parseICalendar(rawICalendar, defaultTimezone, href, extractText(response.properties.getetag) || undefined);

	return parsed.events;
}

async function runCalendarQuery(
	context: CalDavContext,
	calendarHref: string,
	body: string,
	defaultTimezone: string,
): Promise<EventInfo[]> {
	const response = await calDavRequest<string>(
		context,
		{
			url: calendarHref,
			body,
			headers: {
				Depth: '1',
				'Content-Type': 'application/xml; charset=utf-8',
			},
			json: false,
		},
		'REPORT',
	);

	return parseXmlResponses(response).flatMap((entry) =>
		mapEventFromResponse(defaultTimezone, calendarHref, entry),
	);
}

export async function queryEvents(
	context: CalDavContext,
	calendarHref: string,
	options: EventQueryOptions,
): Promise<EventInfo[]> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;
	const defaultTimezone = getCredentialsTimezone(credentials);
	const report = buildCalendarQueryReport({
		start: options.start,
		end: options.end,
		expand: options.expand,
	});

	let events = await runCalendarQuery(context, calendarHref, report, defaultTimezone);

	if (options.filters?.uid) {
		const uidMatches = events.filter((event) => event.uid === options.filters?.uid);
		if (uidMatches.length > 0) {
			events = uidMatches;
		} else {
			const uidReport = buildUidQueryReport(options.filters.uid);
			events = await runCalendarQuery(context, calendarHref, uidReport, defaultTimezone);
		}
	}

	events = events.filter((event) => eventMatchesFilters(event, options.filters));
	events.sort((left, right) => left.start.localeCompare(right.start));

	if (options.limit && options.limit > 0) {
		return events.slice(0, options.limit);
	}

	return events;
}

export async function getEvent(
	context: CalDavContext,
	calendarHref: string,
	identifier: string,
): Promise<EventInfo> {
	const credentials = (await context.getCredentials('calDavApi')) as unknown as CalDavCredentials;
	const defaultTimezone = getCredentialsTimezone(credentials);
	const trimmedIdentifier = identifier.trim();

	if (trimmedIdentifier.length === 0) {
		throw new Error('An event identifier is required.');
	}

	if (isAbsoluteUrl(trimmedIdentifier) || trimmedIdentifier.includes('/') || trimmedIdentifier.endsWith('.ics')) {
		const eventUrl = isAbsoluteUrl(trimmedIdentifier)
			? trimmedIdentifier
			: toAbsoluteUrl(calendarHref, trimmedIdentifier);
		const response = await calDavRequest<{ body?: string; headers?: Record<string, string> }>(
			context,
			{
				url: eventUrl,
				headers: {
					Accept: 'text/calendar',
				},
				returnFullResponse: true,
				encoding: 'text',
			},
			'GET',
		);
		const rawICalendar = response.body ?? '';
		const parsed = parseICalendar(rawICalendar, defaultTimezone, eventUrl, response.headers?.etag);
		const event = parsed.events[0];
		if (!event) {
			throw new Error(`No VEVENT component was found at ${eventUrl}.`);
		}

		return event;
	}

	const matches = await queryEvents(context, calendarHref, {
		filters: {
			uid: trimmedIdentifier,
		},
		limit: 2,
	});
	if (matches.length === 0) {
		throw new Error(`No event with UID "${identifier}" was found.`);
	}

	return matches[0];
}

export async function createEvent(
	context: CalDavContext,
	calendarHref: string,
	options: {
		rawICalendar?: string;
		event: StructuredEventInput;
		filename?: string;
	},
): Promise<EventInfo> {
	const rawICalendar = options.rawICalendar ?? buildEventCalendar(options.event);
	const filename = getFilenameForEvent(options.event.uid, options.filename);
	const eventUrl = toAbsoluteUrl(calendarHref, filename);

	await calDavRequest(context, {
		url: eventUrl,
		body: rawICalendar,
		headers: {
			'Content-Type': 'text/calendar; charset=utf-8',
			'If-None-Match': '*',
		},
		json: false,
	}, 'PUT');

	return await getEvent(context, calendarHref, eventUrl);
}

export async function updateEvent(
	context: CalDavContext,
	calendarHref: string,
	options: {
		identifier: string;
		rawICalendar?: string;
		event?: StructuredEventInput;
		ifMatchEtag?: string;
	},
): Promise<EventInfo> {
	const existingEvent = await getEvent(context, calendarHref, options.identifier);
	const payload =
		options.rawICalendar ??
		(options.event
			? replaceEventProperties(
					existingEvent.rawICalendar,
					new Map<string, string[]>([
						[
							'UID',
							[serializeCalendarProperty('UID', options.event.uid)],
						],
						[
							'DTSTAMP',
							[
								serializeCalendarProperty(
									'DTSTAMP',
									DateTime.utc().toFormat("yyyyLLdd'T'HHmmss'Z'"),
								),
							],
						],
						[
							'DTSTART',
							[foldIcalLine(`DTSTART${formatIcalDateTime(options.event.start, options.event.allDay, options.event.timezone)}`)],
						],
						[
							'DTEND',
							options.event.end
								? [foldIcalLine(`DTEND${formatIcalDateTime(options.event.end, options.event.allDay, options.event.timezone)}`)]
								: [],
						],
						[
							'DURATION',
							options.event.duration ? [serializeCalendarProperty('DURATION', options.event.duration.toISO())] : [],
						],
						[
							'SUMMARY',
							[serializeCalendarProperty('SUMMARY', escapeIcalText(options.event.summary))],
						],
						[
							'DESCRIPTION',
							options.event.description
								? [serializeCalendarProperty('DESCRIPTION', escapeIcalText(options.event.description))]
								: [],
						],
						[
							'LOCATION',
							options.event.location
								? [serializeCalendarProperty('LOCATION', escapeIcalText(options.event.location))]
								: [],
						],
						[
							'STATUS',
							options.event.status ? [serializeCalendarProperty('STATUS', options.event.status.toUpperCase())] : [],
						],
						[
							'TRANSP',
							options.event.transparency
								? [serializeCalendarProperty('TRANSP', options.event.transparency.toUpperCase())]
								: [],
						],
						[
							'ATTENDEE',
							options.event.attendees.map((attendee) =>
								serializeCalendarProperty('ATTENDEE', `mailto:${attendee}`),
							),
						],
					]),
			  )
			: existingEvent.rawICalendar);

	await calDavRequest(context, {
		url: existingEvent.href,
		body: payload,
		headers: {
			'Content-Type': 'text/calendar; charset=utf-8',
			...(options.ifMatchEtag || existingEvent.etag
				? { 'If-Match': options.ifMatchEtag || existingEvent.etag || '' }
				: {}),
		},
		json: false,
	}, 'PUT');

	return await getEvent(context, calendarHref, existingEvent.href);
}

export async function deleteEvent(
	context: CalDavContext,
	calendarHref: string,
	identifier: string,
	ifMatchEtag?: string,
): Promise<EventInfo> {
	const existingEvent = await getEvent(context, calendarHref, identifier);

	await calDavRequest(context, {
		url: existingEvent.href,
		headers: {
			...(ifMatchEtag || existingEvent.etag ? { 'If-Match': ifMatchEtag || existingEvent.etag || '' } : {}),
		},
	}, 'DELETE');

	return existingEvent;
}

function clipInterval(
	start: DateTime,
	end: DateTime,
	window: TimeWindow,
): { start: DateTime; end: DateTime } | undefined {
	const clippedStart = start < window.start ? window.start : start;
	const clippedEnd = end > window.end ? window.end : end;

	if (clippedEnd <= clippedStart) {
		return undefined;
	}

	return {
		start: clippedStart,
		end: clippedEnd,
	};
}

function mergeIntervals(intervals: Array<{ start: DateTime; end: DateTime }>): Array<{ start: DateTime; end: DateTime }> {
	if (intervals.length === 0) {
		return [];
	}

	const sorted = [...intervals].sort((left, right) => left.start.toMillis() - right.start.toMillis());
	const merged: Array<{ start: DateTime; end: DateTime }> = [sorted[0]];

	for (const interval of sorted.slice(1)) {
		const previous = merged[merged.length - 1];
		if (interval.start <= previous.end) {
			previous.end = interval.end > previous.end ? interval.end : previous.end;
			continue;
		}

		merged.push({ ...interval });
	}

	return merged;
}

export function computeFreeBusy(
	events: EventInfo[],
	window: TimeWindow,
	options: FreeBusyOptions,
): FreeBusyResult {
	const busyIntervals = mergeIntervals(
		events
			.filter((event) => {
				if ((event.status ?? '').toUpperCase() === 'CANCELLED') {
					return false;
				}

				if ((event.transparency ?? '').toUpperCase() === 'TRANSPARENT') {
					return false;
				}

				if ((event.status ?? '').toUpperCase() === 'TENTATIVE' && !options.includeTentativeAsBusy) {
					return false;
				}

				return true;
			})
			.map((event) => {
				const start = DateTime.fromISO(event.start, { zone: event.timezone ?? window.zone });
				const end = event.end
					? DateTime.fromISO(event.end, { zone: event.timezone ?? window.zone })
					: start.plus({ minutes: DEFAULT_EVENT_DURATION_MINUTES });

				return clipInterval(start, end, window);
			})
			.filter((interval): interval is { start: DateTime; end: DateTime } => interval !== undefined),
	);

	const freeSlots: FreeSlot[] = [];
	let cursor = window.start;

	for (const busyInterval of busyIntervals) {
		if (busyInterval.start > cursor) {
			freeSlots.push(...splitFreeInterval(cursor, busyInterval.start, options.slotMinutes));
		}

		cursor = busyInterval.end > cursor ? busyInterval.end : cursor;
	}

	if (cursor < window.end) {
		freeSlots.push(...splitFreeInterval(cursor, window.end, options.slotMinutes));
	}

	const filteredFreeSlots = freeSlots.filter(
		(slot) => slot.durationMinutes >= options.minimumDurationMinutes,
	);
	const canFitRequestedDuration = filteredFreeSlots.length > 0;
	const durationLabel =
		options.minimumDurationMinutes > 0 ? `${options.minimumDurationMinutes}-minute` : 'requested';
	const humanSummary = canFitRequestedDuration
		? `Found ${filteredFreeSlots.length} free ${durationLabel} slot(s) between ${formatHumanDateTime(
				window.start.toISO() ?? '',
				window.zone,
				false,
		  )} and ${formatHumanDateTime(window.end.toISO() ?? '', window.zone, false)}.`
		: `No free ${durationLabel} slots were found between ${formatHumanDateTime(
				window.start.toISO() ?? '',
				window.zone,
				false,
		  )} and ${formatHumanDateTime(window.end.toISO() ?? '', window.zone, false)}.`;

	return {
		windowStart: window.start.toISO() ?? '',
		windowEnd: window.end.toISO() ?? '',
		busy: busyIntervals.map((interval) => ({
			start: interval.start.toISO() ?? '',
			end: interval.end.toISO() ?? '',
		})),
		free: filteredFreeSlots,
		canFitRequestedDuration,
		humanSummary,
	};
}

function splitFreeInterval(start: DateTime, end: DateTime, slotMinutes?: number): FreeSlot[] {
	const durationMinutes = Math.max(0, Math.round(end.diff(start, 'minutes').minutes));

	if (!slotMinutes || slotMinutes <= 0) {
		return [
			{
				start: start.toISO() ?? '',
				end: end.toISO() ?? '',
				durationMinutes,
			},
		];
	}

	const slots: FreeSlot[] = [];
	let cursor = start;

	while (cursor.plus({ minutes: slotMinutes }) <= end) {
		const slotEnd = cursor.plus({ minutes: slotMinutes });
		slots.push({
			start: cursor.toISO() ?? '',
			end: slotEnd.toISO() ?? '',
			durationMinutes: slotMinutes,
		});
		cursor = slotEnd;
	}

	return slots;
}

export function getTimezone(
	context: CalDavContext,
	itemIndex: number,
	propertyName = 'timezone',
): string {
	const rawValue = context.getNodeParameter(propertyName, itemIndex, DEFAULT_TIMEZONE) as string;
	return rawValue.trim() || DEFAULT_TIMEZONE;
}

export function getPositiveLimit(context: IExecuteFunctions, itemIndex: number, returnAllProperty = 'returnAll', limitProperty = 'limit'): number | undefined {
	const returnAll = context.getNodeParameter(returnAllProperty, itemIndex, true) as boolean;
	if (returnAll) {
		return undefined;
	}

	return context.getNodeParameter(limitProperty, itemIndex, DEFAULT_RANGE_LIMIT) as number;
}

export function getPositiveInteger(context: IExecuteFunctions, itemIndex: number, propertyName: string, fallbackValue: number): number {
	const value = context.getNodeParameter(propertyName, itemIndex, fallbackValue) as number;

	return value > 0 ? Math.round(value) : fallbackValue;
}

export function getStringValue(context: IExecuteFunctions, itemIndex: number, propertyName: string): string | undefined {
	const value = context.getNodeParameter(propertyName, itemIndex, '') as string;
	return value.trim().length > 0 ? value.trim() : undefined;
}

export function getBooleanValue(context: IExecuteFunctions, itemIndex: number, propertyName: string, fallbackValue = false): boolean {
	return context.getNodeParameter(propertyName, itemIndex, fallbackValue) as boolean;
}

export function assertNonEmpty(value: string | undefined, fieldName: string): string {
	if (!value || value.trim().length === 0) {
		throw new Error(`${fieldName} is required.`);
	}

	return value.trim();
}

export function buildNoResultsSummary(resourceName: string, details: string): IDataObject {
	return {
		count: 0,
		results: [],
		humanSummary: `No ${resourceName} matched ${details}.`,
	};
}
