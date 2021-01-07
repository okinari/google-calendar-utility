const PROP_KEY_PREFIX_SYNC_TOKEN: string = 'syncToken_'

export class GoogleCalendarUtility {

	private _calendarId: string = null

	private _properties: GoogleAppsScript.Properties.Properties = PropertiesService.getScriptProperties()

	private _nextSyncToken: string = null

	private _eventList: GoogleAppsScript.Calendar.Schema.Event[] = null

	constructor(calendarId: string,) {
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
		let nextSyncToken: string = null
		while (nextSyncToken === null) {
			let events: GoogleAppsScript.Calendar.Schema.Events = Calendar.Events.list(this._calendarId, options)
			if ('nextSyncToken' in events) {
				nextSyncToken = events.nextSyncToken
				break
			}
			options.pageToken = events.nextPageToken
		}

		this._setNextSyncToken(nextSyncToken)
	}

	private _getNextSyncToken(): string {
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

	public get eventList(): GoogleAppsScript.Calendar.Schema.Event[] {
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

	// イベント参加者と主催者として、特定のユーザを追加する
	public addToEventAttendeesAndOrganizers(event: GoogleAppsScript.Calendar.Schema.Event, email: string, name: string, ): void {

		// 削除予定のイベントの場合、処理しない
		if (event.status === 'cancelled') {
			return
		}
		
		// 新規追加処理が完了している場合、処理しない
		if (event.extendedProperties?.private?.isCompletedAddToEventAttendeesAndOrganizers === 'ok') {
			return
		}

		// 参加者に自分が存在しない場合、追加する
		if (('attendees' in event) === false) {
			event.attendees = []
		}
		let isExistAttendees: boolean = false
		for (let i in event.attendees) {
			if (event.attendees[i].email === email) {
				isExistAttendees = true
				break
			}
		}
		if (isExistAttendees === false) {
			event.attendees.push({
				email: email,
				displayName: name,
				responseStatus: 'accepted',
			})
		}

		// 主催者が存在しない場合、自分を設定
		if (('organizer' in event) === false) {
			event.organizer = {
				email: email,
				displayName: name,
			}
		}

		// 追加処理の完了を知らせるフラグを立てる
		if (('extendedProperties' in event) === false) {
			event.extendedProperties = {}
		}
		if (('private' in event.extendedProperties) === false) {
			event.extendedProperties.private = {}
		}
		event.extendedProperties.private.isCompletedAddToEventAttendeesAndOrganizers = 'ok'
		
		Calendar.Events.update(event, this._calendarId, event.id)
	}
}