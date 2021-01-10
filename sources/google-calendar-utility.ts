import { GoogleCalendarEventUtility } from './google-calendar-event-utility'

const PROP_KEY_PREFIX_SYNC_TOKEN: string = 'syncToken_'

export class GoogleCalendarUtility {

	private _calendarId: string

	private _properties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties()

	private _nextSyncToken: string | null = null

	private _eventList: GoogleAppsScript.Calendar.Schema.Event[] | null = null

	constructor(calendarId: string, ) {
		this._calendarId = calendarId
	}

	/**
	 * このアプリを使う前に、一度だけ行う処理
	 */
	public initializeGoogleCalendarSync(): void {

		// nextSyncTokenがプロパティが設定されていない場合、現在の状態でのtokenを取得
		this._getNewNextSyncToken()
	}

	/**
	 * nextSyncTokenを新規に取得して、プロパティに保存する
	 */
	private _getNewNextSyncToken(): void {

		// nextSyncTokenがプロパティが設定されている場合、再度の取得処理はしない
		if (this._nextSyncToken !== null) {
			return
		}

		let tmpDate: Date = new Date()
		let timeMax: string = tmpDate.toISOString()
		tmpDate.setDate(tmpDate.getDate() - 1)
		let timeMin: string = tmpDate.toISOString()
		let options: {
			maxResults: number,
			timeMin: string,
			timeMax: string,
			pageToken?: string,
		} = {
			maxResults: 2500,
			timeMin: timeMin,
			timeMax: timeMax,
		}
		let nextSyncToken: string | null = null
		while (nextSyncToken === null) {
			let events: GoogleAppsScript.Calendar.Schema.Events = Calendar.Events.list(this._calendarId, options)
			if ('nextSyncToken' in events && events.nextSyncToken) {
				nextSyncToken = events.nextSyncToken
				break
			}
			options.pageToken = events.nextPageToken
		}

		this._setNextSyncToken(nextSyncToken)
	}

	private _getNextSyncToken(): string | null {
		if (this._nextSyncToken === null) {
			this._nextSyncToken = this._properties.getProperty(PROP_KEY_PREFIX_SYNC_TOKEN + this._calendarId)
			if (this._nextSyncToken === null) {
				this._getNewNextSyncToken()
			}
		}
		return this._nextSyncToken
	}

	private _setNextSyncToken(nextSyncToken: string, ) {
		this._properties.setProperty(PROP_KEY_PREFIX_SYNC_TOKEN + this._calendarId, nextSyncToken)
	}

	public get eventList(): GoogleAppsScript.Calendar.Schema.Event[] | null {
		if (this._eventList === null) {
			let events = Calendar.Events.list(this._calendarId, {
				syncToken: this._getNextSyncToken(),
				maxResults: 2500,
			})
			this._eventList = events.items
			this._setNextSyncToken(events.nextSyncToken)
		}
		return this._eventList
	}

	/**
	 * 引数に与えたイベントを同期する
	 * 同期先のイベントには、同期元のイベントIDを設定しておき、同期元のイベントが更新されたとき、同期先も再更新できるようにする
	 * @param event イベント
	 * @param calendarIdOfOriginal 同期元のカレンダーID
	 * @param calendarIdOfSyncTo 同期先のカレンダーID
	 * @param isPrivate 秘匿化するかどうか
	 */
	public static syncEvent(
			event: GoogleAppsScript.Calendar.Schema.Event,
			calendarIdOfOriginal: string,
			calendarIdOfSyncTo: string,
			options: {
				isPrivate: boolean,
				onlyAttendeesAndOrganizer: {
					email: string,
					name: string,
				} | null,
			} = {
				isPrivate: false,
				onlyAttendeesAndOrganizer: null,
			},
	): void {
		
		// 削除予定のイベントの場合、処理しない
		// 同期して削除したいが、拡張プロパティの値が取得できないので諦めた
		if (event.status === 'cancelled') {
			return
		}

		// この処理にて生成されたイベント(つまり、同期元のイベントIDが設定されているイベント)の場合、処理しない
		else if (event.extendedProperties?.private?.eventIdOfOriginal) {
			return
		}

		// 既にこの処理が実行されているイベント(つまり、同期先のイベントIDが設定されているイベント)の場合、同期先のイベントを更新する
		else if (event.extendedProperties?.private?.eventIdOfSyncTo) {
			// オプション処理をする
			let updateEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarUtility._syncEventOptionsProcess(
				event,
				event.id,
				options,
			)
			Calendar.Events.update(updateEvent, calendarIdOfSyncTo, event.extendedProperties.private.eventIdOfSyncTo)
		}

		// 既にこの処理が実行されているイベント(つまり、同期先のイベントIDが設定されているイベント)の場合、同期先のイベントを更新する
		// 以前、eventIdOfSyncToではなく、eventIdOfCopyToという項目名で書き込んでいたことがあるので、両方とも確認
		else if (event.extendedProperties?.private?.eventIdOfCopyTo) {
			// オプション処理をする
			let updateEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarUtility._syncEventOptionsProcess(
				event,
				event.id,
				options,
			)
			Calendar.Events.update(updateEvent, calendarIdOfSyncTo, event.extendedProperties.private.eventIdOfCopyTo)
		}

		// まだ実行されていない場合、新規追加する
		else {
			// 同期元となるイベントIDを保持して、同期元のイベントIDを削除
			let eventIdOfOriginal: string = event.id
			event.id = null
			
			// オプション処理をする
			let newEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarUtility._syncEventOptionsProcess(
				event,
				eventIdOfOriginal,
				options,
			)

			// イベントを登録
			newEvent = Calendar.Events.insert(newEvent, calendarIdOfSyncTo)

			// 同期元のイベントに対して、同期先のイベントIDを設定して、更新する
			event = GoogleCalendarEventUtility.addPrivateExtendedProperties(event, 'eventIdOfSyncTo', newEvent.id)
			Calendar.Events.update(event, calendarIdOfOriginal, eventIdOfOriginal)
		}
	}

	private static _syncEventOptionsProcess(
		event: GoogleAppsScript.Calendar.Schema.Event,
		eventIdOfOriginal: string,
		options: {
			isPrivate: boolean,
			onlyAttendeesAndOrganizer: {
				email: string,
				name: string,
			} | null
		},
	): GoogleAppsScript.Calendar.Schema.Event {

		let retEvent: GoogleAppsScript.Calendar.Schema.Event = event

		// 同期先イベントには、同期元のイベントIDを設定する
		retEvent = GoogleCalendarEventUtility.addPrivateExtendedProperties(
			retEvent,
			'eventIdOfOriginal',
			eventIdOfOriginal
		)

		// 全てを秘匿化する
		if (options.isPrivate) {
			retEvent = GoogleCalendarEventUtility.convertPrivateEvent(retEvent)
		}
		// 特定のユーザのみを参加者および主催者とする
		else if (options.onlyAttendeesAndOrganizer) {
			retEvent = GoogleCalendarEventUtility.deleteToEventAttendeesAndOrganizers(retEvent)
			retEvent = GoogleCalendarEventUtility.addToEventAttendeesAndOrganizers(
				retEvent,
				options.onlyAttendeesAndOrganizer.email,
				options.onlyAttendeesAndOrganizer.name,
			)
		}

		return retEvent
	}

	/**
	 * 引数に与えたイベントをバックアップする
	 * バックアップ先のイベントには、バックアップ元のイベントIDを設定しておくが、バックアップ元のイベントが更新されても、バックアップ先は再更新されない
	 * @param event イベント
	 * @param calendarIdOfBackupTo バックアップ先のカレンダーID
	 */
	public static backupEvent(
			event: GoogleAppsScript.Calendar.Schema.Event,
			calendarIdOfBackupTo: string,
	): void {
		
		// 削除予定のイベントの場合、処理しない
		// 同期して削除したいが、拡張プロパティの値が取得できないので諦めた
		if (event.status === 'cancelled') {
			return
		}

		// この処理、もしくは同期処理にて生成されたイベント(つまり、バックアップ元のイベントIDが設定されているイベント)の場合、処理しない
		else if (event.extendedProperties?.private?.eventIdOfOriginal) {
			return
		}

		// まだ実行されていない場合、新規追加する
		else {
			let insertEvent: GoogleAppsScript.Calendar.Schema.Event = event

			// 念のため、オリジナル関数でコピーする
			insertEvent = GoogleCalendarEventUtility.copyEvent(insertEvent)

			// バックアップ元となるイベントIDを設定
			insertEvent = GoogleCalendarEventUtility.addPrivateExtendedProperties(insertEvent, 'eventIdOfOriginal', insertEvent.id)

			// バックアップ用のイベントなので、参加者と主催者を全て削除する
			insertEvent = GoogleCalendarEventUtility.deleteToEventAttendeesAndOrganizers(insertEvent)

			// 元々のイベントIDは削除する
			insertEvent.id = null 

			// 新規イベントを登録する
			Calendar.Events.insert(insertEvent, calendarIdOfBackupTo)
		}
	}

	/**
	 * privateな拡張プロパティに項目を追加して、更新する
	 * @param event 
	 * @param itemName 
	 * @param itemValue 
	 */
	public updateToAddPrivateExtendedProperties(event: GoogleAppsScript.Calendar.Schema.Event, itemName: string, itemValue: string, ): void {
		let targetEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarEventUtility.addPrivateExtendedProperties(event, itemName, itemValue)
		Calendar.Events.update(targetEvent, this._calendarId, event.id)
	}
}