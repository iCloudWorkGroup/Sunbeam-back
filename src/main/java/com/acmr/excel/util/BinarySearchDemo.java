package com.acmr.excel.util;

public class BinarySearchDemo {
	int[] array = null;
	int count = 0; // 比较的次数

	public void init() {
		array = new int[] { 0, 2, 4, 6, 8, 9, 11, 13, 15 };
	}

	/**
	 * 普通for循环查询指定的数字例子
	 * 
	 * @param num
	 * @return
	 */
	public int forExample(int num) {
		System.out.println("****************使用普通循环查找****************");
		for (int i = 0; i < array.length; i++) {
			count++;
			if (array.length - 1 == i && num != array[i]) {
				System.out.println("抱歉，没有找到");
			} else if (num == array[i]) {
				System.out.println(array[i] + "找到了，在数组下标为" + i + "的地方，查找了"
						+ count + "次。");
				break;
			}
		}
		return count;
	}

	/**
	 * 二分法查询指定数字 的例子
	 * 
	 * @return
	 */
	public int binarySearch(int num) {
		System.out.println("****************使用二分法查找****************");
		int index = 0; // 检索的时候
		int start = 0; // 用start和end两个索引控制它的查询范围
		int end = array.length - 1;
		int result = -1;
		// count = 0;
		for (int i = 0; i < array.length; i++) {
			// count++;
			index = (start + end) / 2;
			if (array.length - 1 == i) {
				// System.out.println(num + "找到了，在数组下标为" + end + "的地方,查找了" +
				// count + "次。");
				result = end;
				break;
			} else if (array[index] < num) {
				if (array[index + 1] > num) {
					// System.out.println(array[index] + "找到了，在数组下标为" + index +
					// "的地方,查找了" + count + "次。");
					result = index;
					break;
				}
				start = index;
			} else if (array[index] > num) {
				end = index;
			} else {
				// System.out.println(array[index] + "找到了，在数组下标为" + index +
				// "的地方,查找了" + count + "次。");
				result = index;
				break;
			}
		}
		return result;
	}

	// public static void main(String[] args){
	// BinarySearchDemo demo = new BinarySearchDemo();
	// demo.init();
	// int num = 14;
	// //demo.forExample(num);
	// int result = demo.binarySearch(num);
	// System.out.println(result);
	// }
}
