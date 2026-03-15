require "test_helper"

class VbaToolControllerTest < ActionDispatch::IntegrationTest
  test "should get index" do
    get vba_tool_index_url
    assert_response :success
  end

  test "should get manual" do
    get vba_tool_manual_url
    assert_response :success
  end

  test "should get examples" do
    get vba_tool_examples_url
    assert_response :success
  end
end
