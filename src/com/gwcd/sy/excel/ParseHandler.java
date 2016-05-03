package com.gwcd.sy.excel;

import java.util.HashMap;
import java.util.Map;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

class ParseHandler extends DefaultHandler {
	private static final String Q_NAME_STR = "string";
	private static final String ATTR_NAME = "name";
	private Map<String, String> mStringMap = new HashMap<>();
	private IParseListener mParseListener;
	private String mCurName = "";
	private boolean isValidStr = false;

	public void setParseListener(IParseListener parseListener) {
		mParseListener = parseListener;
	}

	@Override
	public void startDocument() throws SAXException {
		super.startDocument();
		if (mParseListener != null) {
			mParseListener.onStart();
		}
	}

	@Override
	public void startElement(String uri, String localName, String qName, Attributes attributes) throws SAXException {
		super.startElement(uri, localName, qName, attributes);
		if (Q_NAME_STR.equals(qName)) {
			isValidStr = true;
		}
		// ¥¶¿Ì Ù–‘
		for (int i = 0; i < attributes.getLength(); ++i) {
			String attrName = attributes.getQName(i);
			String attrValue = attributes.getValue(i);
			if (ATTR_NAME.equals(attrName)) {
				mCurName = attrValue;
			}
		}
	}

	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		String value = new String(ch, start, length);
		if (value != null && !value.isEmpty() && isValidStr && !mCurName.isEmpty()) {
			mStringMap.put(mCurName, value);
		}
	}

	@Override
	public void endElement(String uri, String localName, String qName) throws SAXException {
		super.endElement(uri, localName, qName);
		isValidStr = false;
		mCurName = "";
	}

	@Override
	public void endDocument() throws SAXException {
		super.endDocument();
		System.out.println(mStringMap.toString());
		if (!mStringMap.isEmpty()) {

		}
		if (mParseListener != null) {
			mParseListener.onFinish(mStringMap);
		}
	}
}
