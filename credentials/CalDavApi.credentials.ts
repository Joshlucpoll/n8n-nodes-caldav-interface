import type {
	ICredentialDataDecryptedObject,
	ICredentialTestRequest,
	ICredentialType,
	IHttpRequestOptions,
	Icon,
	INodeProperties,
} from 'n8n-workflow';

export class CalDavApi implements ICredentialType {
	name = 'calDavApi';

	displayName = 'CalDAV API';

	icon: Icon = { light: 'file:../nodes/CalDav/caldav.svg', dark: 'file:../nodes/CalDav/caldav.dark.svg' };

	documentationUrl = 'https://github.com/joshlucpoll/n8n-nodes-caldav-interface#credentials';

	supportedNodes = ['n8n-nodes-caldav-interface.calDav'];

	properties: INodeProperties[] = [
		{
			displayName: 'Base URL',
			name: 'baseUrl',
			type: 'string',
			required: true,
			default: '',
			placeholder: 'https://calendar.example.com',
			description: 'Base URL for the CalDAV server or account root',
		},
		{
			displayName: 'Username',
			name: 'username',
			type: 'string',
			required: true,
			default: '',
			description: 'Username for HTTP Basic authentication',
		},
		{
			displayName: 'Password',
			name: 'password',
			type: 'string',
			required: true,
			typeOptions: {
				password: true,
			},
			default: '',
			description: 'Password for HTTP Basic authentication',
		},
		{
			displayName: 'Calendar Home Path',
			name: 'calendarHomePath',
			type: 'string',
			default: '',
			placeholder: '/calendars/user/',
			description:
				'Optional explicit calendar-home path. Use this when the server does not expose principal discovery cleanly.',
		},
		{
			displayName: 'Default Timezone',
			name: 'defaultTimezone',
			type: 'string',
			default: 'UTC',
			description:
				'Timezone used when the node must interpret natural language dates or floating iCalendar timestamps',
		},
		{
			displayName: 'Ignore TLS Errors',
			name: 'ignoreTlsErrors',
			type: 'boolean',
			default: false,
			description:
				'Whether to ignore TLS certificate errors. Enable only for local development or trusted self-signed servers',
		},
	];

	test: ICredentialTestRequest = {
		request: {
			baseURL: '={{$credentials?.baseUrl}}',
			url: '/',
			headers: {
				Authorization:
					'={{"Basic " + Buffer.from($credentials.username + ":" + $credentials.password).toString("base64")}}',
			},
		},
	};

	async authenticate(
		credentials: ICredentialDataDecryptedObject,
		requestOptions: IHttpRequestOptions,
	): Promise<IHttpRequestOptions> {
		requestOptions.headers ??= {};
		requestOptions.auth = {
			username: credentials.username as string,
			password: credentials.password as string,
		};

		return requestOptions;
	}
}
