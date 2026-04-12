declare module 'luxon' {
	export class DateTime {
		isValid: boolean;
		weekday: number;
		year: number;
		static fromFormat(value: string, format: string, options?: Record<string, unknown>): DateTime;
		static fromISO(value: string, options?: Record<string, unknown>): DateTime;
		static fromJSDate(value: Date, options?: Record<string, unknown>): DateTime;
		static now(): DateTime;
		static utc(): DateTime;
		startOf(unit: string): DateTime;
		endOf(unit: string): DateTime;
		setZone(zone: string): DateTime;
		set(values: Record<string, number>): DateTime;
		toUTC(): DateTime;
		toFormat(format: string): string;
		toISO(): string | null;
		toMillis(): number;
		plus(values: Record<string, number> | Duration): DateTime;
		diff(other: DateTime, unit?: string): Duration;
	}

	export class Duration {
		isValid: boolean;
		minutes: number;
		static fromISO(value: string): Duration;
		static fromObject(values: Record<string, number>): Duration;
		toISO(): string;
	}
}
