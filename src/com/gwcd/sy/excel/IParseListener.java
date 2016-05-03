package com.gwcd.sy.excel;

import java.util.Map;

public interface IParseListener {

	void onStart();
	void onFinish(Map<String, String> stringMap);
}
