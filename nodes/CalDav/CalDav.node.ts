import type {
	IExecuteFunctions,
	IDataObject,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodeListSearchResult,
	INodeProperties,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import {
	NodeApiError,
	NodeConnectionTypes,
	NodeOperationError,
} from 'n8n-workflow';

import {
	assertNonEmpty,
	buildNoResultsSummary,
	computeFreeBusy,
	createCalendar,
	createEvent,
	deleteCalendar,
	deleteEvent,
	describeEvent,
	formatHumanDateTime,
	getBooleanValue,
	getCalendar,
	getEvent,
	getPositiveInteger,
	getPositiveLimit,
	getStringValue,
	getTimezone,
	listCalendars,
	queryEvents,
	resolveCalendarHref,
	resolveEventInput,
	resolveTimeWindow,
	searchCalendars,
	updateEvent,
} from './GenericFunctions';

const resourceProperty: INodeProperties = {
	displayName: 'Resource',
	name: 'resource',
	type: 'options',
	noDataExpression: true,
	default: 'calendar',
	options: [
		{
			name: 'Calendar',
			value: 'calendar',
		},
		{
			name: 'Event',
			value: 'event',
		},
		{
			name: 'Query',
			value: 'query',
		},
	],
};

const calendarOperationProperty: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	default: 'list',
	displayOptions: {
		show: {
			resource: ['calendar'],
		},
	},
	options: [
		{
			name: 'List',
			value: 'list',
			action: 'List calendars',
			description: 'List available calendars',
		},
		{
			name: 'Create',
			value: 'create',
			action: 'Create a calendar',
			description: 'Create a new calendar collection',
		},
		{
			name: 'Delete',
			value: 'delete',
			action: 'Delete a calendar',
			description: 'Delete a calendar collection',
		},
		{
			name: 'Get',
			value: 'get',
			action: 'Get calendar properties',
			description: 'Read calendar metadata and CalDAV properties',
		},
	],
};

const eventOperationProperty: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	default: 'create',
	displayOptions: {
		show: {
			resource: ['event'],
		},
	},
	options: [
		{
			name: 'Create',
			value: 'create',
			action: 'Create an event',
			description: 'Create a calendar event',
		},
		{
			name: 'Get',
			value: 'get',
			action: 'Get event details',
			description: 'Fetch one event or a date-range view of events',
		},
		{
			name: 'Update',
			value: 'update',
			action: 'Update an event',
			description: 'Update an existing calendar event',
		},
		{
			name: 'Delete',
			value: 'delete',
			action: 'Delete an event',
			description: 'Delete an existing calendar event',
		},
	],
};

const queryOperationProperty: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	default: 'filter',
	displayOptions: {
		show: {
			resource: ['query'],
		},
	},
	options: [
		{
			name: 'Filter Events',
			value: 'filter',
			action: 'Query events with filters',
			description: 'Query events with date and text filters',
		},
		{
			name: 'Get Free/Busy',
			value: 'freeBusy',
			action: 'Get free or busy information',
			description: 'Compute busy blocks and available slots in a time range',
		},
	],
};

const calendarLocatorProperty: INodeProperties = {
	displayName: 'Calendar',
	name: 'calendar',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description:
		'@agentic Choose the calendar to work with. You can select one from the server, paste a calendar URL, or enter a collection path such as work or team/project.',
	displayOptions: {
		show: {
			resource: ['calendar', 'event', 'query'],
			operation: ['get', 'delete', 'create', 'get', 'update', 'delete', 'filter', 'freeBusy'],
		},
		hide: {
			resource: ['calendar'],
			operation: ['list', 'create'],
		},
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			typeOptions: {
				searchListMethod: 'searchCalendars',
				searchable: true,
			},
		},
		{
			displayName: 'Calendar URL',
			name: 'url',
			type: 'string',
			placeholder: 'https://calendar.example.com/calendars/user/work/',
		},
		{
			displayName: 'Collection Path',
			name: 'id',
			type: 'string',
			placeholder: 'work',
		},
	],
};

function createRangeFields(displayOptions: INodeProperties['displayOptions']): INodeProperties[] {
	return [
	{
		displayName: 'Range Start',
		name: 'rangeStart',
		type: 'dateTime',
		default: '',
		description:
			'Start of the date range to search. Leave blank and use Natural Language Range if you prefer relative dates such as tomorrow or next week.',
		displayOptions,
	},
	{
		displayName: 'Range End',
		name: 'rangeEnd',
		type: 'dateTime',
		default: '',
		description:
			'End of the date range to search. Leave blank only when Natural Language Range already defines both bounds.',
		displayOptions,
	},
	{
		displayName: 'Natural Language Range',
		name: 'rangeText',
		type: 'string',
		default: '',
		placeholder: 'Friday afternoon',
		description:
			'@agentic Optional natural language window to search, such as tomorrow, Friday afternoon, next week, or 2026-04-15 09:00 to 2026-04-15 17:00',
		displayOptions,
	},
	{
		displayName: 'Expand Recurrence',
		name: 'expandRecurrence',
		type: 'boolean',
		default: true,
		description: 'Whether to expand recurring events into individual occurrences inside the requested range when the CalDAV server supports it',
		displayOptions,
	},
];
}

const eventDateRangeFields = createRangeFields({
	show: {
		resource: ['event'],
		operation: ['get'],
		eventGetMode: ['dateRange'],
	},
});

const queryRangeFields = createRangeFields({
	show: {
		resource: ['query'],
		operation: ['filter', 'freeBusy'],
	},
});

const nodeProperties: INodeProperties[] = [
	resourceProperty,
	calendarOperationProperty,
	eventOperationProperty,
	queryOperationProperty,
	{
		displayName: 'Calendar Path',
		name: 'calendarPath',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'work',
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['create'],
			},
		},
		description: '@agentic Relative path for the new calendar collection, such as work, personal, or team/releases',
	},
	{
		displayName: 'Calendar Name',
		name: 'calendarName',
		type: 'string',
		required: true,
		default: '',
		placeholder: 'Work',
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['create'],
			},
		},
		description: '@agentic Human-readable name for the new calendar, such as Work, Personal, or Sales Team',
	},
	{
		displayName: 'Calendar Description',
		name: 'calendarDescription',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['calendar'],
				operation: ['create'],
			},
		},
		description: '@agentic Optional note describing the calendar purpose, ownership, or expected event type',
	},
	calendarLocatorProperty,
	{
		displayName: 'Get Mode',
		name: 'eventGetMode',
		type: 'options',
		noDataExpression: true,
		default: 'single',
		options: [
			{
				name: 'Single Event',
				value: 'single',
			},
			{
				name: 'Date Range',
				value: 'dateRange',
			},
		],
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['get'],
			},
		},
		description: '@agentic Choose whether to fetch one event by identifier or return all events in a date range',
	},
	{
		displayName: 'Event Identifier',
		name: 'eventIdentifier',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['get', 'update', 'delete'],
				eventGetMode: ['single'],
			},
		},
		description:
			'@agentic Event identifier. You can pass a full event URL, a filename like meeting.ics, or an iCalendar UID.',
	},
	{
		displayName: 'Title',
		name: 'createTitle',
		type: 'string',
		default: '',
		required: true,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description: '@agentic Short human-readable event title, such as Team Sync, Dentist Appointment, or Project Kickoff',
	},
	{
		displayName: 'Updated Title',
		name: 'updateTitle',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['update'],
			},
		},
		description:
			'@agentic Optional new title for the event. Leave blank to keep the current title.',
	},
	{
		displayName: 'When',
		name: 'whenText',
		type: 'string',
		default: '',
		placeholder: 'next Tuesday at 3pm for 45 minutes',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: '@agentic Natural language schedule for the event, such as next Tuesday at 3pm for 45 minutes, April 15th at 10am, or 2026-04-15 09:00 to 2026-04-15 10:00',
	},
	{
		displayName: 'Start',
		name: 'start',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description:
			'@agentic Optional explicit start time. Use this when you want precise scheduling instead of the natural language When field.',
	},
	{
		displayName: 'End',
		name: 'end',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description:
			'@agentic Optional explicit end time. If omitted, the node uses Duration or a default one-hour event.',
	},
	{
		displayName: 'Duration',
		name: 'duration',
		type: 'string',
		default: '',
		placeholder: '45m',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description:
			'@agentic Optional duration like 30m, 2h, or PT45M. This is useful when you know how long the event should last but not the exact end time.',
	},
	{
		displayName: 'All Day',
		name: 'allDay',
		type: 'boolean',
		default: false,
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: 'Whether to treat the event as an all-day event instead of a timed meeting',
	},
	{
		displayName: 'Timezone',
		name: 'timezone',
		type: 'string',
		default: 'UTC',
		displayOptions: {
			show: {
				resource: ['event', 'query'],
				operation: ['create', 'update', 'get', 'filter', 'freeBusy'],
			},
		},
		description: '@agentic Timezone used to interpret natural language dates and times, such as Europe/London or America/New_York',
	},
	{
		displayName: 'Description',
		name: 'eventDescription',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: '@agentic Optional event notes, agenda, or context to store in the calendar description field',
	},
	{
		displayName: 'Location',
		name: 'location',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: '@agentic Optional event location, room, address, or meeting link',
	},
	{
		displayName: 'Status',
		name: 'status',
		type: 'options',
		default: '',
		options: [
			{ name: 'Keep Existing / Unset', value: '' },
			{ name: 'Confirmed', value: 'CONFIRMED' },
			{ name: 'Tentative', value: 'TENTATIVE' },
			{ name: 'Cancelled', value: 'CANCELLED' },
		],
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: '@agentic Optional scheduling status for the event',
	},
	{
		displayName: 'Transparency',
		name: 'transparency',
		type: 'options',
		default: '',
		options: [
			{ name: 'Keep Existing / Unset', value: '' },
			{ name: 'Busy', value: 'OPAQUE' },
			{ name: 'Free', value: 'TRANSPARENT' },
		],
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description:
			'@agentic Control whether the event blocks free/busy time. Busy blocks time, Free does not.',
	},
	{
		displayName: 'Attendees',
		name: 'attendeesText',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create', 'update'],
			},
		},
		description: '@agentic Optional attendee email addresses separated by commas, semicolons, or new lines',
	},
	{
		displayName: 'Custom UID',
		name: 'customUid',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description:
			'@agentic Optional custom iCalendar UID. Leave blank to let the node generate one.',
	},
	{
		displayName: 'Filename',
		name: 'filename',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['create'],
			},
		},
		description:
			'@agentic Optional CalDAV resource filename, such as team-sync.ics. Leave blank to use the UID as the filename.',
	},
	{
		displayName: 'If-Match ETag',
		name: 'etag',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['event'],
				operation: ['update', 'delete'],
			},
		},
		description:
			'@agentic Optional ETag for optimistic concurrency control. Provide this to avoid overwriting someone else’s changes.',
	},
	...eventDateRangeFields,
	...queryRangeFields,
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		default: true,
		displayOptions: {
			show: {
				resource: ['event', 'query'],
				operation: ['get', 'filter'],
			},
			hide: {
				resource: ['event'],
				eventGetMode: ['single'],
			},
		},
		description: 'Whether to return all results or only up to a given limit',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		typeOptions: {
			minValue: 1,
			numberPrecision: 0,
		},
		default: 50,
		displayOptions: {
			show: {
				resource: ['event', 'query'],
				operation: ['get', 'filter'],
				returnAll: [false],
			},
			hide: {
				resource: ['event'],
				eventGetMode: ['single'],
			},
		},
		description: 'Max number of results to return',
	},
	{
		displayName: 'Summary Contains',
		name: 'summaryContains',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Only return events whose titles contain this text',
	},
	{
		displayName: 'Description Contains',
		name: 'descriptionContains',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Only return events whose descriptions contain this text',
	},
	{
		displayName: 'Location Contains',
		name: 'locationContains',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Only return events whose location contains this text',
	},
	{
		displayName: 'UID',
		name: 'filterUid',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Exact iCalendar UID to match when you want a precise event lookup',
	},
	{
		displayName: 'Attendee Email',
		name: 'attendeeEmail',
		type: 'string',
		default: '',
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Only return events containing an attendee email address that matches this value',
	},
	{
		displayName: 'Filter Status',
		name: 'filterStatus',
		type: 'options',
		default: '',
		options: [
			{ name: 'Any', value: '' },
			{ name: 'Confirmed', value: 'CONFIRMED' },
			{ name: 'Tentative', value: 'TENTATIVE' },
			{ name: 'Cancelled', value: 'CANCELLED' },
		],
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['filter'],
			},
		},
		description: '@agentic Only return events with this scheduling status',
	},
	{
		displayName: 'Minimum Duration Minutes',
		name: 'minimumDurationMinutes',
		type: 'number',
		typeOptions: {
			minValue: 1,
			numberPrecision: 0,
		},
		default: 60,
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['freeBusy'],
			},
		},
		description: '@agentic Minimum meeting length, in minutes, that must fit inside a free slot',
	},
	{
		displayName: 'Slot Minutes',
		name: 'slotMinutes',
		type: 'number',
		typeOptions: {
			minValue: 1,
			numberPrecision: 0,
		},
		default: 30,
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['freeBusy'],
			},
		},
		description: '@agentic Optional slot size used to break larger free blocks into bookable chunks',
	},
	{
		displayName: 'Count Tentative As Busy',
		name: 'includeTentativeAsBusy',
		type: 'boolean',
		default: true,
		displayOptions: {
			show: {
				resource: ['query'],
				operation: ['freeBusy'],
			},
		},
		description: 'Whether to treat tentative events as busy time when computing availability',
	},
];

function mapCalendarResult(calendar: Awaited<ReturnType<typeof getCalendar>>, action: string) {
	return {
		...calendar,
		humanSummary:
			action === 'list'
				? `Calendar "${calendar.displayName}" at ${calendar.href}.`
				: `${action === 'create' ? 'Created' : action === 'delete' ? 'Deleted' : 'Loaded'} calendar "${calendar.displayName}" at ${calendar.href}.`,
	};
}

function mapEventResult(event: Awaited<ReturnType<typeof getEvent>>, action: string) {
	return {
		...event,
		humanSummary:
			action === 'delete'
				? `Deleted ${describeEvent(event)}.`
				: action === 'create'
					? `Created ${describeEvent(event)}.`
					: action === 'update'
						? `Updated ${describeEvent(event)}.`
						: `Loaded ${describeEvent(event)}.`,
	};
}

export class CalDav implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'CalDAV Interface',
		name: 'calDav',
		icon: { light: 'file:caldav.svg', dark: 'file:caldav.dark.svg' },
		group: ['output'],
		version: [1],
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Manage calendars and events on CalDAV servers',
		defaults: {
			name: 'CalDAV Interface',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		credentials: [
			{
				name: 'calDavApi',
				required: true,
			},
		],
		properties: nodeProperties,
	};

	methods = {
		listSearch: {
			async searchCalendars(this: ILoadOptionsFunctions, filter?: string): Promise<INodeListSearchResult> {
				try {
					return await searchCalendars(this, filter);
				} catch {
					return { results: [] };
				}
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const resource = this.getNodeParameter('resource', itemIndex) as string;
				const operation = this.getNodeParameter('operation', itemIndex) as string;

				if (resource === 'calendar') {
					if (operation === 'list') {
						const calendars = await listCalendars(this);

						if (calendars.length === 0) {
							returnData.push({
								json: buildNoResultsSummary('calendars', 'the current credentials') as IDataObject,
								pairedItem: { item: itemIndex },
							});
							continue;
						}

						for (const calendar of calendars) {
							returnData.push({
								json: mapCalendarResult(calendar, 'list') as IDataObject,
								pairedItem: { item: itemIndex },
							});
						}

						continue;
					}

					if (operation === 'create') {
						const calendar = await createCalendar(this, {
							path: assertNonEmpty(getStringValue(this, itemIndex, 'calendarPath'), 'Calendar Path'),
							displayName: assertNonEmpty(getStringValue(this, itemIndex, 'calendarName'), 'Calendar Name'),
							description: getStringValue(this, itemIndex, 'calendarDescription'),
						});

						returnData.push({
							json: mapCalendarResult(calendar, 'create') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}

					if (operation === 'get') {
						const calendar = await getCalendar(
							this,
							this.getNodeParameter('calendar', itemIndex) as { mode?: string; value?: string },
						);

						returnData.push({
							json: mapCalendarResult(calendar, 'get') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}

					if (operation === 'delete') {
						const calendar = await getCalendar(
							this,
							this.getNodeParameter('calendar', itemIndex) as { mode?: string; value?: string },
						);
						await deleteCalendar(
							this,
							this.getNodeParameter('calendar', itemIndex) as { mode?: string; value?: string },
						);

						returnData.push({
							json: mapCalendarResult(calendar, 'delete') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}
				}

				if (resource === 'event') {
					const calendarHref = await resolveCalendarHref(
						this,
						this.getNodeParameter('calendar', itemIndex) as { mode?: string; value?: string },
					);
					const timezone = getTimezone(this, itemIndex);

					if (operation === 'create') {
						const event = resolveEventInput({
							title: assertNonEmpty(getStringValue(this, itemIndex, 'createTitle'), 'Title'),
							whenText: getStringValue(this, itemIndex, 'whenText'),
							start: getStringValue(this, itemIndex, 'start'),
							end: getStringValue(this, itemIndex, 'end'),
							duration: getStringValue(this, itemIndex, 'duration'),
							allDay: getBooleanValue(this, itemIndex, 'allDay', false),
							timezone,
							description: getStringValue(this, itemIndex, 'eventDescription'),
							location: getStringValue(this, itemIndex, 'location'),
							status: getStringValue(this, itemIndex, 'status'),
							transparency: getStringValue(this, itemIndex, 'transparency'),
							uid: getStringValue(this, itemIndex, 'customUid'),
							attendeesText: getStringValue(this, itemIndex, 'attendeesText'),
						});

						const createdEvent = await createEvent(this, calendarHref, {
							event,
							filename: getStringValue(this, itemIndex, 'filename'),
						});

						returnData.push({
							json: mapEventResult(createdEvent, 'create') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}

					if (operation === 'get') {
						const eventGetMode = this.getNodeParameter('eventGetMode', itemIndex, 'single') as string;

						if (eventGetMode === 'single') {
							const event = await getEvent(
								this,
								calendarHref,
								assertNonEmpty(getStringValue(this, itemIndex, 'eventIdentifier'), 'Event Identifier'),
							);

							returnData.push({
								json: mapEventResult(event, 'get') as IDataObject,
								pairedItem: { item: itemIndex },
							});
							continue;
						}

						const window = resolveTimeWindow(
							getStringValue(this, itemIndex, 'rangeText'),
							getStringValue(this, itemIndex, 'rangeStart'),
							getStringValue(this, itemIndex, 'rangeEnd'),
							timezone,
						);
						const events = await queryEvents(this, calendarHref, {
							start: window.start,
							end: window.end,
							expand: getBooleanValue(this, itemIndex, 'expandRecurrence', true),
							limit: getPositiveLimit(this, itemIndex),
						});

						if (events.length === 0) {
							returnData.push({
								json: buildNoResultsSummary(
									'events',
									`${formatHumanDateTime(window.start.toISO() ?? '', timezone)} and ${formatHumanDateTime(window.end.toISO() ?? '', timezone)}`,
								) as IDataObject,
								pairedItem: { item: itemIndex },
							});
							continue;
						}

						for (const event of events) {
							returnData.push({
								json: {
									...mapEventResult(event, 'get'),
									querySummary: `Found ${events.length} event(s) between ${formatHumanDateTime(
										window.start.toISO() ?? '',
										timezone,
									)} and ${formatHumanDateTime(window.end.toISO() ?? '', timezone)}.`,
								} as IDataObject,
								pairedItem: { item: itemIndex },
							});
						}
						continue;
					}

					if (operation === 'update') {
						const existingEvent = await getEvent(
							this,
							calendarHref,
							assertNonEmpty(getStringValue(this, itemIndex, 'eventIdentifier'), 'Event Identifier'),
						);
						const title = getStringValue(this, itemIndex, 'updateTitle') ?? existingEvent.summary;
						const updatedInput = resolveEventInput({
							title,
							whenText: getStringValue(this, itemIndex, 'whenText'),
							start: getStringValue(this, itemIndex, 'start'),
							end: getStringValue(this, itemIndex, 'end'),
							duration: getStringValue(this, itemIndex, 'duration'),
							allDay: getBooleanValue(this, itemIndex, 'allDay', existingEvent.allDay),
							timezone,
							description: getStringValue(this, itemIndex, 'eventDescription'),
							location: getStringValue(this, itemIndex, 'location'),
							status: getStringValue(this, itemIndex, 'status'),
							transparency: getStringValue(this, itemIndex, 'transparency'),
							attendeesText: getStringValue(this, itemIndex, 'attendeesText'),
							existingEvent,
						});
						const updatedEvent = await updateEvent(this, calendarHref, {
							identifier: existingEvent.href,
							event: updatedInput,
							ifMatchEtag: getStringValue(this, itemIndex, 'etag'),
						});

						returnData.push({
							json: mapEventResult(updatedEvent, 'update') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}

					if (operation === 'delete') {
						const deletedEvent = await deleteEvent(
							this,
							calendarHref,
							assertNonEmpty(getStringValue(this, itemIndex, 'eventIdentifier'), 'Event Identifier'),
							getStringValue(this, itemIndex, 'etag'),
						);

						returnData.push({
							json: mapEventResult(deletedEvent, 'delete') as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}
				}

				if (resource === 'query') {
					const calendarHref = await resolveCalendarHref(
						this,
						this.getNodeParameter('calendar', itemIndex) as { mode?: string; value?: string },
					);
					const timezone = getTimezone(this, itemIndex);
					const window = resolveTimeWindow(
						getStringValue(this, itemIndex, 'rangeText'),
						getStringValue(this, itemIndex, 'rangeStart'),
						getStringValue(this, itemIndex, 'rangeEnd'),
						timezone,
					);

					if (operation === 'filter') {
						const events = await queryEvents(this, calendarHref, {
							start: window.start,
							end: window.end,
							expand: getBooleanValue(this, itemIndex, 'expandRecurrence', true),
							limit: getPositiveLimit(this, itemIndex),
							filters: {
								uid: getStringValue(this, itemIndex, 'filterUid'),
								summaryContains: getStringValue(this, itemIndex, 'summaryContains'),
								descriptionContains: getStringValue(this, itemIndex, 'descriptionContains'),
								locationContains: getStringValue(this, itemIndex, 'locationContains'),
								status: getStringValue(this, itemIndex, 'filterStatus'),
								attendeeEmail: getStringValue(this, itemIndex, 'attendeeEmail'),
							},
						});

						if (events.length === 0) {
							returnData.push({
								json: buildNoResultsSummary(
									'events',
									`the supplied filters between ${formatHumanDateTime(window.start.toISO() ?? '', timezone)} and ${formatHumanDateTime(window.end.toISO() ?? '', timezone)}`,
								) as IDataObject,
								pairedItem: { item: itemIndex },
							});
							continue;
						}

						for (const event of events) {
							returnData.push({
								json: {
									...mapEventResult(event, 'get'),
									querySummary: `Found ${events.length} matching event(s) between ${formatHumanDateTime(
										window.start.toISO() ?? '',
										timezone,
									)} and ${formatHumanDateTime(window.end.toISO() ?? '', timezone)}.`,
								} as IDataObject,
								pairedItem: { item: itemIndex },
							});
						}
						continue;
					}

					if (operation === 'freeBusy') {
						const events = await queryEvents(this, calendarHref, {
							start: window.start,
							end: window.end,
							expand: true,
						});
						const freeBusy = computeFreeBusy(events, window, {
							includeTentativeAsBusy: getBooleanValue(this, itemIndex, 'includeTentativeAsBusy', true),
							minimumDurationMinutes: getPositiveInteger(
								this,
								itemIndex,
								'minimumDurationMinutes',
								60,
							),
							slotMinutes: getPositiveInteger(this, itemIndex, 'slotMinutes', 30),
						});

						returnData.push({
							json: {
								...freeBusy,
								matchedEvents: events.map((event) => ({
									uid: event.uid,
									summary: event.summary,
									start: event.start,
									end: event.end,
									status: event.status,
									transparency: event.transparency,
								})),
							} as IDataObject,
							pairedItem: { item: itemIndex },
						});
						continue;
					}
				}

				throw new NodeOperationError(this.getNode(), `Unsupported operation ${resource}.${operation}.`, {
					itemIndex,
				});
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({
						json: {
							error: (error as Error).message,
						},
						pairedItem: { item: itemIndex },
					});
					continue;
				}

				if (
					error instanceof Error &&
					('httpCode' in error || 'response' in error || 'statusCode' in error)
				) {
					throw new NodeApiError(
						this.getNode(),
						{ message: (error as Error).message },
						{ itemIndex },
					);
				}

				throw new NodeOperationError(this.getNode(), error as Error, { itemIndex });
			}
		}

		return [returnData];
	}
}
