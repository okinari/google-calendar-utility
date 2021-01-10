export namespace GoogleCalendarEventUtility {

	/**
	 * イベントをコピーしたものを返却する
	 * @param event イベント
	 */
	export function copyEvent(
			event: GoogleAppsScript.Calendar.Schema.Event,
	): GoogleAppsScript.Calendar.Schema.Event {
		let retEvent: GoogleAppsScript.Calendar.Schema.Event = {}
		for (let key in event) {
			if (event[key]) {
				if (key === 'organizer') {
					retEvent.organizer = {}
					for (let okey in event.organizer) {
						retEvent.organizer[okey] = event.organizer[okey]
					}
				}
				else if (key === 'attendees') {
					retEvent.attendees = []
					for (let i in event.attendees) {
						let tmp: GoogleAppsScript.Calendar.Schema.EventAttendee = {}
						for (let akey in event.attendees[i]) {
							tmp[akey] = event.attendees[i][akey]
						}
						retEvent.attendees.push(tmp)
					}
				}
				else if (key === 'extendedProperties') {
					retEvent.extendedProperties = {}
					if (event.extendedProperties.private) {
						retEvent.extendedProperties.private = {}
						for (let ekey in event.extendedProperties.private) {
							retEvent.extendedProperties.private[ekey] = event.extendedProperties.private[ekey]
						}
					}
					if (event.extendedProperties.shared) {
						retEvent.extendedProperties.shared = {}
						for (let ekey in event.extendedProperties.shared) {
							retEvent.extendedProperties.shared[ekey] = event.extendedProperties.shared[ekey]
						}
					}
				}
				else {
					retEvent[key] = event[key]
				}
			}
		}
		return retEvent
	}

	/**
	 * 特定のユーザを、イベント参加者かつ主催者として追加する
	 * @param event イベント
	 * @param email 追加するユーザのメールアドレス
	 * @param name 追加するユーザの名前
	 */
	export function addToEventAttendeesAndOrganizers(
			event: GoogleAppsScript.Calendar.Schema.Event,
			email: string,
			name: string,
	): GoogleAppsScript.Calendar.Schema.Event {
		let retEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarEventUtility.copyEvent(event)
		// 参加者に自分が存在しない場合、追加する
		if (('attendees' in retEvent) === false) {
			retEvent.attendees = []
		}
		let isExistAttendees: boolean = false
		for (let i in retEvent.attendees) {
			if (retEvent.attendees[i].email === email) {
				isExistAttendees = true
				break
			}
		}
		if (isExistAttendees === false) {
			retEvent.attendees.push({
				email: email,
				displayName: name,
				responseStatus: 'accepted',
			})
		}

		// 主催者が存在しない場合、自分を設定
		if (('organizer' in retEvent) === false) {
			retEvent.organizer = {
				email: email,
				displayName: name,
			}
		}
		return retEvent
	}

	/**
	 * イベント参加者と主催者を全て削除する
	 * @param event イベント
	 */
	export function deleteToEventAttendeesAndOrganizers(
		event: GoogleAppsScript.Calendar.Schema.Event,
	): GoogleAppsScript.Calendar.Schema.Event {
		let retEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarEventUtility.copyEvent(event)
		retEvent.attendees = []
		retEvent.organizer = {}
		return retEvent
	}

	/**
	 * イベントの内容を全て秘匿化する
	 * @param event 
	 */
	export function convertPrivateEvent(
			event: GoogleAppsScript.Calendar.Schema.Event,
	): GoogleAppsScript.Calendar.Schema.Event {
		let retEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarEventUtility.copyEvent(event)
		retEvent.visibility = 'private'
		retEvent.summary = '予定あり'
		retEvent.description = ''
		retEvent.location = ''
		retEvent.organizer = {
			id: '',
			email: '',
			displayName: '',
			self: false,
		}
		retEvent.attendees = []
		return retEvent
	}

	export function addPrivateExtendedProperties(
			event: GoogleAppsScript.Calendar.Schema.Event,
			itemName: string,
			itemValue: string,
	): GoogleAppsScript.Calendar.Schema.Event {
		let retEvent: GoogleAppsScript.Calendar.Schema.Event = GoogleCalendarEventUtility.copyEvent(event)
		if (('extendedProperties' in retEvent) === false) {
			retEvent.extendedProperties = {}
		}
		if (('private' in retEvent.extendedProperties) === false) {
			retEvent.extendedProperties.private = {}
		}
		retEvent.extendedProperties.private[itemName] = itemValue
		return retEvent
	}
}