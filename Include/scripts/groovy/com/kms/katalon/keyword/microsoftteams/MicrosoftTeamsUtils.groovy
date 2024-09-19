package com.kms.katalon.keyword.microsoftteams

import java.util.concurrent.atomic.AtomicInteger

import org.apache.commons.lang3.StringUtils
import org.apache.http.client.methods.HttpPost
import org.apache.http.entity.StringEntity
import org.apache.http.impl.client.CloseableHttpClient
import org.apache.http.impl.client.HttpClients
import org.apache.http.util.EntityUtils

import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.context.TestCaseContext
import com.kms.katalon.core.context.TestSuiteContext
import com.kms.katalon.core.setting.BundleSettingStore
import com.kms.katalon.core.util.KeywordUtil

public class MicrosoftTeamsUtils {


	static BundleSettingStore bundleSetting
	static String URL
	static boolean enabled = false

	Map<String, AtomicInteger> stats = new HashMap<String, AtomicInteger>()

	static {
		try {
			bundleSetting = new BundleSettingStore(RunConfiguration.getProjectDir(), 'MicrosoftTeamsIncomingWebhook', true)
			URL = bundleSetting.getString('MicrosoftTeamsIncomingWebhook', '')
			if (StringUtils.isBlank(URL)) {
				KeywordUtil.markWarning("[MS Team] Microsoft Teams Webhook URL is empty.")
			} else {
				KeywordUtil.logInfo("[MS Team] Microsoft Teams integration is enabled")
				enabled = true
			}
		} catch (Exception e) {
			e.printStackTrace()
		}
	}

	/**
	 * Get test cases status.
	 *
	 * @param testCaseContext related information of the executed test case.
	 */
	public getTestcaseStatus (TestCaseContext testCaseContext) {
		if (enabled) {
			String status = testCaseContext.getTestCaseStatus()
			AtomicInteger stat = stats.get(status)
			if (stat == null) {
				stat = new AtomicInteger(0)
			}
			stat.getAndIncrement()
			stats.put(status, stat)
		}
	}

	/**
	 * Update test result summary to Microsoft Team.
	 *
	 * @param testSuiteContext related information of the executed test suite.
	 */
	public updateMicrosoftTeam(TestSuiteContext testSuiteContext) {
		if (enabled) {
			String message = "Summary execution result of Test Suite: " + testSuiteContext.getTestSuiteId()
			for (Map.Entry<String, AtomicInteger> entry : stats.entrySet()) {
				message =  message + "\n\n${entry.getKey()}: ${entry.getValue()}"
			}
			sendMessageToTeamsChannel(URL, message)
		}
	}

	def sendMessageToTeamsChannel(String webhookUrl, String message) {
		CloseableHttpClient httpClient = HttpClients.createDefault()
		try {
			HttpPost request = new HttpPost(webhookUrl)
			String body = """{
       "type":"message",
       "attachments":[
          {
             "contentType":"application/vnd.microsoft.card.adaptive",
             "contentUrl":null,
             "content":{
                "\$schema":"http://adaptivecards.io/schemas/adaptive-card.json",
                "type":"AdaptiveCard",
                "version":"1.2",
                "body":[
                    {
                    "type": "TextBlock",
                    "text": "${message}",
					"wrap": true
                    }
                ]
             }
          }
       ]
    }"""
			StringEntity params = new StringEntity(body, "UTF-8")
			request.addHeader("Content-Type", "application/json")
			request.setEntity(params)

			def response = httpClient.execute(request)
			def responseBody = EntityUtils.toString(response.getEntity())

			if (response.getStatusLine().getStatusCode() >= 200 && response.getStatusLine().getStatusCode() <= 299) {
				KeywordUtil.logInfo("[MS Team] Message sent successfully!")
			} else {
				KeywordUtil.markWarning("[MS Team] Failed to send message. Status: ${response.getStatusLine().getStatusCode()}")
				KeywordUtil.markWarning("[MS Team] Response: ${responseBody}")
			}
		} finally {
			httpClient.close()
		}
	}
}
