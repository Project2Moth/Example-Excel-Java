package example.config;

import java.util.Locale;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.support.ReloadableResourceBundleMessageSource;
import org.springframework.web.servlet.LocaleResolver;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;
import org.springframework.web.servlet.i18n.LocaleChangeInterceptor;
import org.springframework.web.servlet.i18n.SessionLocaleResolver;

/**
 * The type 18 n configuration.
 */
@Configuration
public class I18nConfiguration implements WebMvcConfigurer {

  @Bean
  public ReloadableResourceBundleMessageSource messageSource() {
    ReloadableResourceBundleMessageSource bms = new ReloadableResourceBundleMessageSource();
    bms.setBasenames("classpath:i18n/messages", "classpath:i18n/core_messages");
    bms.setDefaultEncoding("UTF-8");
    return bms;
  }

  /**
   * Locale resolver locale resolver.
   *
   * @return the locale resolver
   */
  @Bean
  public LocaleResolver localeResolver() {
    SessionLocaleResolver sessionLocaleResolver = new SessionLocaleResolver();
    sessionLocaleResolver.setDefaultLocale(new Locale("vi", "VN"));
    return sessionLocaleResolver;
  }

  /**
   * Locale change interceptor locale change interceptor.
   *
   * @return the locale change interceptor
   */
  @Bean
  public LocaleChangeInterceptor localeChangeInterceptor() {
    LocaleChangeInterceptor localeChangeInterceptor = new LocaleChangeInterceptor();
    localeChangeInterceptor.setParamName("lang");
    return localeChangeInterceptor;
  }

  @Override
  public void addInterceptors(InterceptorRegistry registry) {
    registry.addInterceptor(localeChangeInterceptor());
  }
}
